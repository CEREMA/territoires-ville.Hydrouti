VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_decant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Bassin de Décantation"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "Frm_decant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleMode       =   0  'User
   ScaleWidth      =   9339.302
   Begin VB.TextBox Tb_volume 
      BackColor       =   &H80000016&
      Height          =   1005
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2000
      Width           =   4050
   End
   Begin VB.ComboBox Cb_decant 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   4208
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du bassin de décantation"
      Top             =   3120
      Width           =   1052
   End
   Begin VB.Frame Frm_dec 
      Caption         =   "Paramètres"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.TextBox Tb_dec 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2040
         Width           =   900
      End
      Begin VB.TextBox Tb_dec 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1680
         Width           =   900
      End
      Begin VB.TextBox Tb_dec 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   3720
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Tb_dec 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   2
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Tb_dec 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   1
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Lb_udec 
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   18
         Top             =   1365
         Width           =   400
      End
      Begin VB.Label Lb_udec 
         Caption         =   "m/s"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   15
         Top             =   2085
         Width           =   400
      End
      Begin VB.Label Lb_udec 
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   14
         Top             =   1725
         Width           =   400
      End
      Begin VB.Label Lb_udec 
         Caption         =   "mm"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   13
         Top             =   1005
         Width           =   400
      End
      Begin VB.Label Lb_udec 
         Caption         =   "m3/s"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   12
         Top             =   645
         Width           =   400
      End
      Begin VB.Label Lb_intdec 
         Caption         =   "Vitesse horizontale des particules entre 0.2 et 0.5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2085
         Width           =   3615
      End
      Begin VB.Label Lb_intdec 
         Caption         =   "Pourcentage de sédimentation "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1725
         Width           =   2535
      End
      Begin VB.Label Lb_intdec 
         Caption         =   "Rapport l/h compris entre 1 et 5"
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1365
         Width           =   2415
      End
      Begin VB.Label Lb_intdec 
         Caption         =   "Taille des particules à décanter"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   2415
      End
      Begin VB.Label Lb_intdec 
         Caption         =   "Débit à décanter"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   645
         Width           =   2415
      End
   End
   Begin RichTextLib.RichTextBox tb_resu 
      Height          =   1575
      Left            =   5640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   300
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483626
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frm_decant.frx":08CA
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   6120
      MaxLength       =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   2880
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Bassin de décantation"
      Begin VB.Menu mnunouv 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuouv 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu f1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Enregistrer"
      End
      Begin VB.Menu mnusaves 
         Caption         =   "En&registrer sous..."
      End
      Begin VB.Menu mnusuppr 
         Caption         =   "&Supprimer..."
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuprint 
         Caption         =   "Im&primer..."
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter module"
      End
   End
End
Attribute VB_Name = "Frm_decant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private okg As Boolean
Private owner As MDIFrm_menu
Private esave As st_savchute
Public nom_ouvrage As String
'Private nom_fich As String
Public nom_type As String
Private lhFicDbf As Long
Private FileLength As Integer
Private list_par_text()
Private list_par_val()
Private list_par_unite()
Private list_resu_text()
Private list_resu_val()
Private list_resu_unite()
Private list_don1() As Variant
Private list_int1() As Variant
Private list_resu1() As Variant
Private dec_texte As String
Private fen_titre As String
Public titre_sav As String
Private list_tb() As Variant
Private sval_champ As String
Private iSels As Integer
Private iSell As Integer
Private bKP As Boolean
Private label_prec As String
Private mes_prec As String
Private index_prec As Integer
Private change_coul As Boolean
Private Sub sel_text(tb_objet As TextBox)
    tb_objet.SelStart = 0
    tb_objet.SelLength = Len(tb_objet.Text)
End Sub
Private Sub Change_Couleur(nom As String, Index As Integer)
'Dim coul As ColorConstants, coulp As ColorConstants
'Dim Index1 As Integer
'Dim nom1 As String
'coulp = vbBlack
'coul = Couleur_Change
'nom1 = nom
'Select Case nom
'    Case Is = "Tb_dec"
'         nom1 = "Lb_intdec"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intdec"
'         Lb_intdec(index_prec).ForeColor = coulp
'    Case Is = "Frm_dec"
'         Frm_dec.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intdec"
'         Lb_intdec(Index).ForeColor = coul
'         Tb_dec(Index).SetFocus
'   Case Is = "Frm_dec"
'         Tb_dec(0).SetFocus
'         DoEvents
'         Lb_intdec(0).ForeColor = coulp
'         Frm_dec.ForeColor = coul
'End Select
'label_prec = nom1
'index_prec = Index
'change_coul = True
End Sub
Private Sub Change_Focus(nom As String, Index As Integer)
Dim coul As ColorConstants, coulp As ColorConstants
Dim Index1 As Integer
Dim nom1 As String
coulp = vbBlack
coul = Couleur_Change
nom1 = nom
Select Case nom1
    Case Is = "Me"
         Me.SetFocus
    Case Is = "Lb_intdec"
         Tb_dec(Index).SetFocus
   Case Is = "Frm_dec"
         Tb_dec(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_intdec", "Tb_dec", "Frm_dec"
        mes = IDhlp_DecantationModeCalcul  '"Mode de calcul hydraulique d'un bassin de décantation"
End Select
mes_prec = mes
Rec_Mes = mes
End Function
Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Private Sub init_l_tab()
Dim l0() As Variant ', l1() As Variant, l2() As Variant
l0 = Array(0)
'l1 = Array(0,"TB_car_ep", "TB_car_eu", "TB_carep_rur")
'l2 = Array(0,"TB_par_ep", "TB_par_eu", "TB_par_pl")
ReDim list_tb(0 To UBound(l0)) ', 0 To UBound(l1), 0 To UBound(l2))
list_tb = Array(l0) ' , l1, l2)
End Sub
Public Sub retailler()
retaille
End Sub
Private Sub retaille()
    Me.Left = owner.fcom.Width + owner.fcom.Left
    Me.Top = 0
'    Me.Width = owner.Width - owner.fcom.Width - 200
'    Me.Height = owner.fdessin.Top
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc) ' 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub
Private Sub lect_fich()
Dim za As st_savdecant
Dim za1 As st_savdec1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_decant.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdecant
        If Trim(za.type) = nom_type Then
            Cb_decant.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
dec_texte = Cb_decant.list(0)
Cb_decant.Text = Cb_decant.list(0)
Cb_decant.Refresh

Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Sub Mquitter()
    MnuQuit_Click
End Sub
Public Sub Mquit()
    m_quitter_Click
End Sub
Public Sub Msupprimer()
    mnusuppr_Click
End Sub
Public Sub Menregistrer()
    mnusave_Click
End Sub
Public Sub Mimprimer()
    mnuprint_Click
End Sub
Public Sub Mnouveau()
    mnunouv_Click
End Sub
Public Sub Menregsous()
    mnusaves_Click
End Sub
Public Sub Mouvrir()
    mnuouv_Click
End Sub
Public Sub Minfo()
    mnuinfo_Click
End Sub
Private Sub ini_lbresu()
'    Me.Tb_volume.BackColor = &H8000000B
    Me.Tb_volume.BorderStyle = 1
    Me.Tb_volume.Text = ""
'   Me.tb_resu.BackColor = &H8000000B
    Me.tb_resu.BorderStyle = 1
    Me.tb_resu.Text = ""
End Sub
Private Sub modi_lbresu()
'   Me.tb_resu.BackColor = &H80000009
    Me.tb_resu.BorderStyle = 1
'    Me.Tb_volume.BackColor = &H80000009
    Me.Tb_volume.BorderStyle = 1
End Sub
Private Function recup_Vsed(ByVal Gran As Double) As Double
'houpie 20040122
'Dim list_TC(5, 2) As Double
Dim list_TC(11, 2) As Double
Dim a As Double

Dim i As Integer

'list_TC(0, 1) = 0#
'list_TC(1, 1) = 0.125
'list_TC(2, 1) = 0.16
'list_TC(3, 1) = 0.2
'list_TC(4, 1) = 0.25
'list_TC(5, 1) = 0.315
'list_TC(0, 2) = 0#
'list_TC(1, 2) = 0.86
'list_TC(2, 2) = 1.35
'list_TC(3, 2) = 1.9
'list_TC(4, 2) = 2.55
'list_TC(5, 2) = 3.5
list_TC(0, 1) = 0#
list_TC(1, 1) = 0.005
list_TC(2, 1) = 0.01
list_TC(3, 1) = 0.02
list_TC(4, 1) = 0.05
list_TC(5, 1) = 0.1
list_TC(6, 1) = 0.125
list_TC(7, 1) = 0.16
list_TC(8, 1) = 0.2
list_TC(9, 1) = 0.25
list_TC(10, 1) = 0.315
list_TC(11, 1) = 0.5

list_TC(0, 2) = 0#
list_TC(1, 2) = 0.0018
list_TC(2, 2) = 0.007
list_TC(3, 2) = 0.03
list_TC(4, 2) = 0.19
list_TC(5, 2) = 0.7
list_TC(6, 2) = 0.86
list_TC(7, 2) = 1.35
list_TC(8, 2) = 1.9
list_TC(9, 2) = 2.55
list_TC(10, 2) = 3.5
list_TC(11, 2) = 5.8

i = 0
While Gran > list_TC(i, 1) And i < UBound(list_TC)
    i = i + 1
    
Wend
If i = 0 Then
    i = 1
End If
a = (Gran - list_TC(i - 1, 1)) * (list_TC(i, 2) - list_TC(i - 1, 2)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 2)
recup_Vsed = a
End Function
Private Function recup_majK(ByVal Gran As Double, ByVal Psed As Double) As Double
Dim list_TC(5, 4) As Double
Dim list_Prct(3) As Double
Dim list_TCr(3) As Double
Dim a As Double
Dim k As Double

Dim i As Integer, j As Integer
list_Prct(1) = 100
list_Prct(2) = 90
list_Prct(3) = 85

list_TC(1, 1) = 0.125
list_TC(2, 1) = 0.16
list_TC(3, 1) = 0.2
list_TC(4, 1) = 0.25
list_TC(5, 1) = 0.315
list_TC(1, 2) = 5.06
list_TC(2, 2) = 4.67
list_TC(3, 2) = 4.12
list_TC(4, 2) = 3.45
list_TC(5, 2) = 2.84
list_TC(1, 3) = 3.28
list_TC(2, 3) = 3.07
list_TC(3, 3) = 2.43
list_TC(4, 3) = 2.04
list_TC(5, 3) = 1.75
list_TC(1, 4) = 2.75
list_TC(2, 4) = 2.4
list_TC(3, 4) = 1.92
list_TC(4, 4) = 1.59
list_TC(5, 4) = 1.48



i = 1
While Gran > list_TC(i, 1) And i < UBound(list_TC)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
list_TCr(1) = (Gran - list_TC(i - 1, 1)) * (list_TC(i, 2) - list_TC(i - 1, 2)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 2)
list_TCr(2) = (Gran - list_TC(i - 1, 1)) * (list_TC(i, 3) - list_TC(i - 1, 3)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 3)
list_TCr(3) = (Gran - list_TC(i - 1, 1)) * (list_TC(i, 4) - list_TC(i - 1, 4)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 4)
i = 1
While Psed < list_Prct(i) And i < UBound(list_Prct)
i = i + 1
Wend
k = (Psed - list_Prct(i - 1)) * (list_TCr(i) - list_TCr(i - 1)) / (list_Prct(i) - list_Prct(i - 1)) + list_TCr(i - 1)



recup_majK = k
End Function



Private Sub Cb_decant_Change()
    Cb_decant.Text = dec_texte
End Sub

Private Sub Cb_decant_KeyDown(KeyCode As Integer, Shift As Integer)
    dec_texte = Cb_decant.Text
    Cb_decant.Text = dec_texte

End Sub

Private Sub Cb_decant_KeyPress(KeyAscii As Integer)
    dec_texte = Cb_decant.Text
End Sub

Private Sub Cmd_calcul_Click()
Dim Vsed As Double, majK As Double, Sect As Double
Dim Q As Double, Vh As Double, h As Double, larg As Double, X As Double
Dim Longueur As Double, sresult As String, sresult1 As String
Dim temps As Double
Dim ebvolume As volume_dess
ebvolume.coef = 0.3
Q = ebdecant.Q
Vh = ebdecant.Vhor
X = ebdecant.X
    Vsed = recup_Vsed(ebdecant.d)
    majK = recup_majK(ebdecant.d, ebdecant.Psed)
' calcul section transversale
    Sect = Q / Vh
'calcul hauteur et largeur
    h = Sqr(Sect / X)
    larg = X * h
    temps = h / (Vsed / 100)
    Longueur = majK * Vh * temps
    ebdecant.Long = Round(Longueur, 2)
    ebdecant.larg = Round(larg, 2)
    ebdecant.heau = Round(h, 2)
    ebdecant.Vvert = Round(Vsed, 3)
    ebdecant.k = Round(majK, 2)
    
    Call modi_lbresu
    sresult = "  Longueur de la chambre de décantation = " + ajout_zero(Trim(str(Round(Longueur, 2)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  Largeur de la chambre de décantation = " + ajout_zero(Trim(str(Round(larg, 2)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  Hauteur d'eau à l'entrée du décanteur = " + ajout_zero(Trim(str(Round(h, 2)))) + " m"
    sresult1 = "  Vitesse verticale des particules = " + ajout_zero(Trim(str(Round(Vsed, 3)))) + " cm/s"
    sresult1 = sresult1 + Chr(13) + Chr(10) + "  Facteur de majoration de Kalbskopf = " + ajout_zero(Trim(str(Round(majK, 2))))
    Me.tb_resu.Text = sresult1
    Me.Tb_volume.Text = sresult
    
    
'    Call init_graph(owner.fdessin.UC_graphique1)
'    Call dess_decant(owner.fdessin.UC_graphique1)
    ebvolume.Largeur = ebdecant.larg
    ebvolume.Longueur = ebdecant.Long
    ebvolume.Profondeur = ebdecant.heau
    If ebvolume.Longueur > 0 And ebvolume.Largeur > 0 And ebvolume.Profondeur > 0 Then
        Call init_graph_rect(owner.fdessin.UC_graphique1, ebvolume)
        Call dess_stock_rect(owner.fdessin.UC_graphique1, ebvolume)
        owner.fdessin.UC_graphique1.dess_lign 0, ebvolume.Profondeur, ebvolume.Longueur, 0, couleur.magenta, 1
        Call init_graph_rect(Frm_desprint.UC_graphique1, ebvolume)
        Call dess_stock_rect(Frm_desprint.UC_graphique1, ebvolume)
        Frm_desprint.UC_graphique1.dess_lign 0, ebvolume.Profondeur, ebvolume.Longueur, 0, couleur.magenta, 1
   End If

    ouv_sauve = True
    Me.Cmd_calcul.Enabled = False
    'impression true
    Me.mnuprint.Enabled = True

End Sub
Public Sub dess_decant(ByRef uc_g As UC_graphique)
Dim xam As Double, yam As Double, xav As Double, yav As Double

uc_g.redef_drwidth 2
xam = 0
yam = 0
xav = ebdecant.Long
yav = 0
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = 0
xav = 0
yav = ebdecant.heau + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = ebdecant.heau + 0.3
xav = ebdecant.Long
yav = ebdecant.heau + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebdecant.Long
yam = 0
xav = ebdecant.Long
yav = ebdecant.heau + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebdecant.Long
yam = 0
xav = ebdecant.Long + 2 * ebdecant.larg
yav = 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebdecant.Long
yam = ebdecant.heau + 0.3
xav = ebdecant.Long + 2 * ebdecant.larg
yav = ebdecant.heau + 0.3 + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebdecant.Long + 2 * ebdecant.larg
yam = 0.3
xav = ebdecant.Long + 2 * ebdecant.larg
yav = ebdecant.heau + 0.3 + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = ebdecant.heau + 0.3
xav = 2 * ebdecant.larg
yav = ebdecant.heau + 0.3 + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 2 * ebdecant.larg
yam = ebdecant.heau + 0.3 + 0.3
xav = ebdecant.Long + 2 * ebdecant.larg
yav = ebdecant.heau + 0.3 + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
uc_g.redef_drwidth 1
xam = 0
yam = 0
xav = 2 * ebdecant.larg
yav = 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
xam = 2 * ebdecant.larg
yam = 0.3
xav = 2 * ebdecant.larg
yav = ebdecant.heau + 0.3 + 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
xam = 2 * ebdecant.larg
yam = 0.3
xav = ebdecant.Long + 2 * ebdecant.larg
yav = 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
uc_g.redef_drwidth 1
xam = 0
yam = ebdecant.heau
xav = ebdecant.Long
yav = ebdecant.heau
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1
xam = ebdecant.Long
yam = ebdecant.heau
xav = ebdecant.Long + 2 * ebdecant.larg
yav = ebdecant.heau + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1
xam = 0
yam = ebdecant.heau
xav = ebdecant.Long
yav = 0
uc_g.dess_lign xam, yam, xav, yav, couleur.magenta, 1

uc_g.redef_drwidth 1
'uc_g.dess_coth 0, 0, ebdecant.Long, 0, ebdecant.Long, couleur_noir
'uc_g.dess_cotv 0, 0, 0, ebdecant.Heau, ebdecant.Heau, couleur_noir
'uc_g.dess_cotb 0, ebdecant.Heau + 0.3, 2 * ebdecant.larg, ebdecant.Heau + 0.3 + 0.3, ebdecant.larg, couleur_noir
uc_g.dess_coth_text 0, 0, ebdecant.Long, 0, ajout_zero(Trim(str(ebdecant.Long))) + " m", couleur_noir
uc_g.dess_cotv_texte 0, 0, 0, ebdecant.heau, ajout_zero(Trim(str(ebdecant.heau))) + " m ", couleur_noir
uc_g.dess_cotb_text 0, ebdecant.heau + 0.3, 2 * ebdecant.larg, ebdecant.heau + 0.3 + 0.3, ebdecant.larg, ajout_zero(Trim(str(ebdecant.larg))) + " m  ", couleur_noir
End Sub
Private Sub init_graph(ByRef uc_graph As UC_graphique)
Dim ok As Boolean
Dim ecx As Double
Dim i As Integer
Dim decalx As Double
ok = False
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
    uc_graph.init_title
    uc_graph.init_titleh ""
    uc_graph.init_titleb ""
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
decalx = ebdecant.Long / 5#
If decalx < 3 Then
    decalx = 3
End If
uc_graph.init_MinX -decalx '4#
uc_graph.init_MaxX ebdecant.Long + 3 * ebdecant.larg
uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
uc_graph.init_MaxY ebdecant.heau + 1
uc_graph.init_MinY -0.5
uc_graph.init_EchYn 1
   
End Sub


Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'un bassin de décantation"
    Change_Couleur "Me", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus

End Sub

Private Sub Form_Load()
    okg = True
   Me.KeyPreview = True
   Call ini_tooltip_decant(Me)
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
''''    owner.affich_aide Me.Name, "Décantation"
'    nom_fich = chemin_app + "ouvrages1.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "decantation"
    fen_titre = Me.Caption
    ouv_sauve = False
    save_fich = False
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Cb_decant.Visible = False
'     Tb_titre.Visible = True
'     Cb_decant.Visible = True
    Frm_desprint.Show
    Frm_desprint.Visible = False
    Call debut
End Sub
Private Sub debut0()
    Cb_decant.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
    sval_champ = ""
    bKP = False
   Call init_l_tab
    Me.Tb_dec(0).Text = "0.0"
    Me.Tb_dec(1).Text = "0.0"
    Me.Tb_dec(2).Text = "0"
    Me.Tb_dec(3).Text = "0.0"
    Me.Tb_dec(4).Text = "0.0"
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Call reini_valeurs
    Call ini_ebdecant
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin de décantation n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de décantation")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
        Case Is = 7
            ouv_sauve = False
        Case Is = 2
            Cancel = True
    End Select
End If
 '   Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frm_menu.Enabled = True
    ouv_sauve = False
    Unload Frm_desprint
    Unload owner.fdessin
    owner.recharge_commentaire

End Sub
Public Sub Cb_decant_click()
Dim za As st_savdecant
Dim za1 As st_savdec1
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    dec_texte = Trim(nom_ouvrage)
    Cb_decant.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdecant
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_decant.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebdecant = za.decant
            Call ini_form
            Call reini_valeurs
            If Cmd_calcul.Enabled Then
                Call Cmd_calcul_Click
            End If
            ouv_sauve = False
            save_fich = True
            If fich_lect <> nom_fich Then
                ouv_sauve = True
            End If
        End If
    End If

Loop
Close #lhFicDbf
If fich_lect <> nom_fich Then
    Kill fich_lect
End If
 
Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub


Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub Frm_dec_Click()
Dim mes As String
Dim nom As String
nom = "Frm_dec"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
'Change_Couleur nom, 0
Change_Focus nom, 0

End Sub


Private Sub Lb_intdec_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intdec"
mes = Rec_Mes(nom, Index)
'Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then
        Me.mnusave.Enabled = True
        Me.mnusaves.Enabled = True
        Me.mnusuppr.Enabled = True
'        Me.mnuprint.Enabled = True
    Else
        Me.mnusave.Enabled = False
        Me.mnusaves.Enabled = False
        Me.mnusuppr.Enabled = False
        Me.mnuprint.Enabled = False
   End If
End Sub

Private Sub mnuinfo_Click()
    Frm_saisie.Show 1
End Sub

Private Sub mnunouv_Click()
Dim reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin de décantation n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de décantation")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
            Call debut0
        Case Is = 7
            Call debut0
    End Select
Else
    Call debut0
End If
End Sub

Private Sub mnuouv_Click()
Dim reponse As Integer
Dim frmf As Frm_lectfich
Set frmf = New Frm_lectfich
Dim nom As String
'modif FO   ' If ProtectCheck(2) <> 0 Then End
fich_lect = nom_fich
If nom_fich_edit <> "" Then
    nom = "Etude " + nom_fich_edit
Else
    nom = " Nouvelle étude "
End If
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin de décantation n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de décantation")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_decant.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de décantation "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_decant_click
            End If
        Case Is = 7
'            Cb_decant.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de décantation "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_decant_click
            End If
    End Select
Else
'    Cb_decant.Visible = True
    frmf.Label1.Caption = "Recherche d'un bassin de décantation "
    frmf.Caption = nom
    frmf.Show 1
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_decant_click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer
ReDim list_don1(Tb_dec.count - 1, 3)
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "decant"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE BASSIN de DECANTATION"
FrmPrint.sstitre1 = Frm_dec.Caption
FrmPrint.ssTitre2 = "Résultats intermédiaires"
FrmPrint.ssTitre3 = ""
Frm_imp.Type1 = "decant"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE BASSIN de DECANTATION"
Frm_imp.sstitre1 = Frm_dec.Caption
Frm_imp.ssTitre2 = "Résultats intermédiaires"
Frm_imp.ssTitre3 = ""
For i = 0 To Tb_dec.count - 1
    list_don1(i, 1) = Lb_intdec(i).Caption
    list_don1(i, 2) = Tb_dec(i).Text
    list_don1(i, 3) = Lb_udec(i).Caption
Next
list_int1 = rec_list(tb_resu.Text)
list_resu1 = rec_list(Tb_volume.Text)
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
'FrmPrint.paint_picture pict1
'SavePicture pict1, chemin_app + "dess.bmp"

'FrmPrint.Show
'Form1.Show
End Sub
Public Function lect_list(ByVal nom As String) As Variant
Select Case nom
Case Is = "list_don1"
    lect_list = list_don1
Case Is = "list_int1"
    lect_list = list_int1
Case Is = "list_resu1"
    lect_list = list_resu1
End Select
End Function

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Private Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If
End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_savdecant
Dim za1 As st_savdec1
Dim i As Integer, isave As Integer
Dim reponse As Integer
 
If Trim(Tb_titre.Text) <> "" Then
    Call funlockb
  lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            za = za1.stsavdecant
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un bassin de décantation")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "decantation"
            za.nom = Tb_titre.Text
            za.decant = ebdecant
            za1.stsavdecant = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        
        End If
    Else
        za.type = "decantation"
        za.nom = Tb_titre.Text
        za.decant = ebdecant
        za1.stsavdecant = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        dec_texte = Trim(Tb_titre.Text)
        Cb_decant.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom du bassin de décantation n'est pas renseigné.", , "Sauvegarde d'un bassin de décantation")
End If
 
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub

Private Sub mnusaves_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'un bassin de décantation "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du bassin de décantation " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du bassin de décantation (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub

Private Sub mnusuppr_Click()
Dim za As st_savdecant
Dim za1 As st_savdec1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_decant.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_decant.Text) + " va être supprimé .", 4, "Suppression d'un bassin de décantation")
    If reponse = 6 Then  '6=oui,7=non
    save_fich = True
    nom = chemin_app + "tempbas.bin"
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
    Do While Not EOF(lhFicDbf)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf, , za1
       If Not EOF(lhFicDbf) Then
            za = za1.stsavdecant
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_decant.Text)) Then
                FileLength = LOF(lhFicDbf1) / Len(za1) + 1
                Put #lhFicDbf1, FileLength, za1
            End If
       End If
    Loop
    Close #lhFicDbf
    Close #lhFicDbf1
    Kill nom_fich
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Write Lock Read Write As #lhFicDbf Len = Len(za1)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Read As #lhFicDbf1 Len = Len(za1)
    Do While Not EOF(lhFicDbf1)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf1, , za1
       If Not EOF(lhFicDbf1) Then
            FileLength = LOF(lhFicDbf) / Len(za1) + 1
            Put #lhFicDbf, FileLength, za1
       End If
    Loop
    Close #lhFicDbf
    Call flockb(nom_fich)
    Close #lhFicDbf1
    Kill nom
    Call lect_fich
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
    Call reini_valeurs
    Call ini_ebdecant
    Me.Tb_dec(0).Text = "0.0"
    Me.Tb_dec(1).Text = "0.0"
    Me.Tb_dec(2).Text = "0"
    Me.Tb_dec(3).Text = "0.0"
    Me.Tb_dec(4).Text = "0.0"
    ouv_sauve = False
    save_fich = False
    End If
End If
 

Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Private Sub reini_valeurs()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Call ini_lbresu
                'impression false
            Me.mnuprint.Enabled = False

    If ebdecant.Q > 0 And ebdecant.d > 0 And ebdecant.X > 0 _
        And ebdecant.Psed > 0 And ebdecant.Vhor > 0 Then
        Me.Cmd_calcul.Enabled = True
    Else
        Me.Cmd_calcul.Enabled = False
    End If
    ouv_sauve = True
End Sub
Public Sub ini_ebdecant()
    ebdecant.Q = 0#
    ebdecant.d = 0#
    ebdecant.X = 0
    ebdecant.Psed = 0#
    ebdecant.Vhor = 0#
    ebdecant.Long = 0#
    ebdecant.larg = 0#
    ebdecant.Hchamb = 0#
    ebdecant.heau = 0#
    ebdecant.Vvert = 0#
    ebdecant.k = 0#
End Sub
Private Sub ini_form()
    Me.Tb_dec(0).Text = rempl_virgule(Format(ebdecant.Q, "#0.000"))
    Me.Tb_dec(1).Text = rempl_virgule(Format(ebdecant.d, "#0.000"))
    Me.Tb_dec(2).Text = rempl_virgule(Format(ebdecant.X, "0"))
    Me.Tb_dec(3).Text = rempl_virgule(Format(ebdecant.Psed, "##0.00"))
    Me.Tb_dec(4).Text = rempl_virgule(Format(ebdecant.Vhor, "##0.00"))
End Sub
Private Sub Tb_dec_change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_dec(Index).Text, "Saisie du débit à décanter", "R")
            Case Is = 1
                nom = verif_cart0(Tb_dec(Index).Text, "Saisie de la taille des particules à décanter", "R")
            Case Is = 2
                nom = verif_cart0(Tb_dec(Index).Text, "Saisie du rapport l/h", "I")
            Case Is = 3
                nom = verif_cart0(Tb_dec(Index).Text, "Saisie du pourcentage de sédimentation", "R")
            Case Is = 4
                nom = verif_cart0(Tb_dec(Index).Text, "Saisie de la vitesse horizontale des particules", "R")
        End Select
  If nom = "" Then
      Tb_dec(Index).Text = sval_champ
        Tb_dec(Index).SelStart = iSels
        Tb_dec(Index).SelLength = iSell

  End If
End If
' valeur avant
    Select Case Index
        Case Is = 0
             ebdecant.Q = txtVersNum(Me.Tb_dec(0).Text)
        Case Is = 1
             ebdecant.d = txtVersNum(Me.Tb_dec(1).Text)
        Case Is = 2
             ebdecant.X = txtVersNum(Me.Tb_dec(2).Text)
        Case Is = 3
             ebdecant.Psed = txtVersNum(Me.Tb_dec(3).Text)
        Case Is = 4
             ebdecant.Vhor = txtVersNum(Me.Tb_dec(4).Text)
    End Select
    Call reini_valeurs
    sval_champ = ""
    bKP = False
End Sub

Private Sub Tb_dec_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dec"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
DoEvents
Me.Show
Call sel_text(Me.Tb_dec(Index))
End Sub

Private Sub Tb_dec_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dec"
Call sel_text(Tb_dec(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If
End Sub

Private Sub Tb_dec_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_dec(Index).Text
    iSels = Tb_dec(Index).SelStart
    iSell = Tb_dec(Index).SelLength

'    If Len(Tb_dec(Index).Text) <= Tb_dec(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_dec(Index).Text, KeyAscii, "Saisie du débit à décanter", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_dec(Index).Text, KeyAscii, "Saisie de la taille des particules à décanter", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_dec(Index).Text, KeyAscii, "Saisie du rapport l/h", "I")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_dec(Index).Text, KeyAscii, "Saisie du pourcentage de sédimentation", "R")
'            Case Is = 4
'                KeyAscii = verif_car(Tb_dec(Index).Text, KeyAscii, "Saisie de la vitesse horizontale des particules", "R")
'        End Select
'    End If
End If
End Sub
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, "" 'Dimensionnement d'un bassin de décantation"
End Sub
Private Sub Tb_dec_LostFocus(Index As Integer)
Dim ok As Boolean
On Error GoTo ErrorHandler
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_dec", Index, txtVersNum(Tb_dec(Index).Text))
    If Not ok Then
         DoEvents
       Tb_dec(Index).SetFocus
    End If
         DoEvents
    okg = True
End If
Exit Sub
ErrorHandler:
    okg = True


End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub

