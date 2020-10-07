VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_siphon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Siphon"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_siphon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.Frame Frm_Aval 
      Caption         =   "Conduite Aval "
      Height          =   1815
      Left            =   5280
      TabIndex        =   19
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Cmd_ava 
         Caption         =   "Courbe..."
         Height          =   255
         Left            =   3210
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Courbe de débit de la conduite aval"
         Top             =   1440
         Width           =   990
      End
      Begin VB.TextBox Tb_ava 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox Tb_ava 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   7
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Tb_ava 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Tb_ava 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   5
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Lb_uava 
         Height          =   255
         Index           =   2
         Left            =   2655
         TabIndex        =   42
         Top             =   1010
         Width           =   495
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Cote radier ZRav"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Lb_uava 
         Caption         =   "m"
         Height          =   255
         Index           =   3
         Left            =   2655
         TabIndex        =   36
         Top             =   1485
         Width           =   300
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Coeff. de  Strickler"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1005
         Width           =   1455
      End
      Begin VB.Label Lb_uava 
         Caption         =   "1/10000"
         Height          =   255
         Index           =   1
         Left            =   2655
         TabIndex        =   23
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Lb_uava 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   2655
         TabIndex        =   22
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Pente"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Diamètre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   285
      Left            =   5280
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin RichTextLib.RichTextBox Tb_resu 
      Height          =   1575
      Left            =   5280
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1980
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483626
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frm_siphon.frx":08CA
   End
   Begin VB.ComboBox Cb_siphon 
      Height          =   315
      Left            =   240
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   4000
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du siphon"
      Top             =   3650
      Width           =   1000
   End
   Begin VB.Frame Frm_siphon 
      Caption         =   "Siphon "
      Height          =   2130
      Left            =   240
      TabIndex        =   24
      Top             =   1850
      Width           =   4695
      Begin VB.CommandButton Cmd_Kc 
         Caption         =   "Saisie coudes..."
         Height          =   300
         Left            =   3240
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Saisie des coudes"
         Top             =   1730
         Width           =   1365
      End
      Begin VB.TextBox Tb_siph 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1720
         Width           =   900
      End
      Begin VB.TextBox Tb_siph 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   12
         Top             =   1320
         Width           =   900
      End
      Begin VB.TextBox Tb_siph 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   11
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Tb_siph 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   10
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Tb_siph 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   9
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Lb_usiph 
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   44
         Top             =   1845
         Width           =   105
      End
      Begin VB.Label Lb_usiph 
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   43
         Top             =   650
         Width           =   495
      End
      Begin VB.Label Lb_intsiph 
         Caption         =   "Coefficient singularités"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   1750
         Width           =   1695
      End
      Begin VB.Label Lb_usiph 
         Caption         =   "m"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   31
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label Lb_usiph 
         Caption         =   "m3/s"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   30
         Top             =   1005
         Width           =   555
      End
      Begin VB.Label Lb_usiph 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   29
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Lb_intsiph 
         Caption         =   "Longueur développée"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   1365
         Width           =   2415
      End
      Begin VB.Label Lb_intsiph 
         Caption         =   "Débit maximum"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   1005
         Width           =   2415
      End
      Begin VB.Label Lb_intsiph 
         Caption         =   "Coeff.  de  Strickler"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   645
         Width           =   2415
      End
      Begin VB.Label Lb_intsiph 
         Caption         =   "Diamètre "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   285
         Width           =   2415
      End
   End
   Begin VB.Frame Frm_Amont 
      Caption         =   "Conduite Amont "
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Cmd_amo 
         Caption         =   "Courbe..."
         Height          =   255
         Left            =   3600
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Courbe de débit de la conduite amont"
         Top             =   1440
         Width           =   990
      End
      Begin VB.TextBox Tb_amo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox Tb_amo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   3
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Tb_amo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   2
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Tb_amo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Lb_uamo 
         Height          =   255
         Index           =   2
         Left            =   2655
         TabIndex        =   41
         Top             =   1010
         Width           =   495
      End
      Begin VB.Label Lb_intamo 
         Caption         =   "Cote radier ZRam"
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Lb_uamo 
         Caption         =   "m"
         Height          =   255
         Index           =   3
         Left            =   2655
         TabIndex        =   34
         Top             =   1485
         Width           =   300
      End
      Begin VB.Label Lb_intamo 
         Caption         =   "Coeff. de  Strickler"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Lb_uamo 
         Caption         =   "1/10000"
         Height          =   255
         Index           =   1
         Left            =   2655
         TabIndex        =   18
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Lb_uamo 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   2655
         TabIndex        =   17
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Lb_intamo 
         Caption         =   "Pente"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lb_intamo 
         Caption         =   "Diamètre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Siphon"
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
Attribute VB_Name = "Frm_siphon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private okg As Boolean
Private owner As MDIFrm_menu
Private esave As st_savsi
Public nom_ouvrage As String
'Private nom_fich As String
Public nom_type As String
Private lhFicDbf As Long
Private FileLength As Integer
Private list_don1() As Variant
Private list_int1() As Variant
Private list_don2() As Variant
Private list_don3() As Variant
Private si_texte As String
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
'    Case Is = "Tb_amo"
'         nom1 = "Lb_intamo"
'    Case Is = "Tb_ava"
'         nom1 = "Lb_intava"
'    Case Is = "Tb_siph"
'         nom1 = "Lb_intsiph"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intamo"
'         Lb_intamo(index_prec).ForeColor = coulp
'    Case Is = "Lb_intava"
'         Lb_intava(index_prec).ForeColor = coulp
'    Case Is = "Lb_intsiph"
'         Lb_intsiph(index_prec).ForeColor = coulp
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coulp
'    Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coulp
'    Case Is = "Frm_siphon"
'         Frm_siphon.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intamo"
'         Lb_intamo(Index).ForeColor = coul
'    Case Is = "Lb_intava"
'         Lb_intava(Index).ForeColor = coul
'    Case Is = "Lb_intsiph"
'         Lb_intsiph(Index).ForeColor = coul
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coul
'   Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coul
'   Case Is = "Frm_siphon"
'         Frm_siphon.ForeColor = coul
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
    Case Is = "Lb_intamo"
         Tb_amo(Index).SetFocus
    Case Is = "Lb_intava"
         Tb_ava(Index).SetFocus
    Case Is = "Lb_intsiph"
         Tb_siph(Index).SetFocus
    Case Is = "Frm_Amont"
         Tb_amo(0).SetFocus
   Case Is = "Frm_Aval"
         Tb_ava(0).SetFocus
   Case Is = "Frm_siphon"
         Tb_siph(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_intamo", "Tb_amo", "Frm_Amont", "Lb_intava", "Tb_ava", "Frm_Aval", "Frm_siphon"
        mes = IDhlp_SiphonPrincipesHydrauliques '"Principes hydrauliques"
    Case Is = "Lb_intsiph", "Tb_siph"
        Select Case Index
            Case Is = 0
                 mes = IDhlp_SiphonPrincipesHydrauliques '"Principes hydrauliques"
            Case Is = 1
                 mes = IDhlp_SiphonPrincipesHydrauliques  '"Principes hydrauliques"
            Case Is = 2
                 mes = IDhlp_SiphonPrincipesHydrauliques  '"Principes hydrauliques"
            Case Is = 3
                 mes = IDhlp_SiphonPrincipesHydrauliques  '"Principes hydrauliques"
            Case Is = 4
                 mes = IDhlp_SiphonPertesChargesSingulieres  '"Pertes de charge singulières dans le siphon (coudes)"
        End Select
    Case Is = "Cmd_Kc"
        mes = IDhlp_SiphonPertesChargesSingulieres  '"Pertes de charge singulières dans le siphon (coudes)"
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

Private Sub Cb_siphon_Change()
    Cb_siphon.Text = si_texte
End Sub

Private Sub Cb_siphon_KeyDown(KeyCode As Integer, Shift As Integer)
    si_texte = Cb_siphon.Text
    Cb_siphon.Text = si_texte
End Sub

Private Sub Cb_siphon_KeyPress(KeyAscii As Integer)
    si_texte = Cb_siphon.Text
End Sub

Private Sub Cmd_amo_Click()
Call dessin_courbe_amo
End Sub

Private Sub Cmd_ava_Click()
Call dessin_courbe_ava
End Sub

Private Sub Cmd_calcul_Click()
    Call calcul_amont_aval
    ouv_sauve = True
End Sub
Private Sub Cmd_Kc_Click()
'Me.Enabled = False
Dim mes As String
Dim nom As String
nom = "Cmd_Kc"
Change_Focus nom, 0
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
DoEvents
Frm_singul.Show 1
End Sub
Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'un siphon"
    Change_Couleur "Me", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Form_Load()
    okg = True
    Me.KeyPreview = True
    Call ini_tooltip_siphon(Me)
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'''''    owner.affich_aide Me.Name, "Siphon"
'    nom_fich = chemin_app + "etude.bin"
'    nom_fich = chemin_app + "siphon.bin"
    nom_type = "siphon"
    fen_titre = Me.Caption
    ouv_sauve = False
    save_fich = False
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Cb_siphon.Visible = False
'        Cmd_amo.DisabledPicture = LoadPicture(chemin_app + "Pavés rouges.bmp")
'        Cmd_amo.Picture = LoadPicture(chemin_app + "Pavés rouges.bmp")
'        Cmd_amo.DownPicture = LoadPicture(chemin_app + "Pavés rouges.bmp")
    Frm_desprint.Show
    Frm_desprint.Visible = False
    Call debut
End Sub
Private Sub debut0()
    Cb_siphon.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
 Call init_l_tab
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
     owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.Height = 4200
   bKP = False
    sval_champ = ""
    Call reini_valeurs
    Call ini_listcoud
    Call ini_ebsiphon
    Call ini_form
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub dessin_courbe_amo()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = ebsiphon.dam / 1000#
    canal.Longueur = 5
    canal.pente = ebsiphon.iRadam / 10000#
    canal.rugosite = ebsiphon.Kam
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = ebsiphon.Rdav
      .radamo = ebsiphon.Rdav + 0.3 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, 0, "Courbe débit conduite amont")
End Sub
Private Sub dessin_courbe_ava()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = ebsiphon.dav / 1000#
    canal.Longueur = 5
    canal.pente = ebsiphon.iradav / 10000#
    canal.rugosite = ebsiphon.kav
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = ebsiphon.Rdam - 0.3
      .radamo = ebsiphon.Rdam 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, 0, "Courbe débit conduite aval")
End Sub
Public Sub dess_siphon(ByRef uc_g As UC_graphique, ByVal h1 As Double, ByVal h2 As Double, ByVal jtot As Double)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim ls As Double, decalx As Double, decaly As Double
ls = ebsiphon.ls * 0.75
decalx = ls / 4
decaly = 0.5
decaly = minimum(ebsiphon.tron_amo.radava, ebsiphon.tron_ava.radamo) - decaly
uc_g.redef_drwidth 2
'dessin conduite amont
xam = ebsiphon.tron_amo.Absamo
yam = ebsiphon.tron_amo.radamo
xav = ebsiphon.tron_amo.Absava - ebsiphon.ds
yav = ebsiphon.tron_amo.radava
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = ebsiphon.tron_amo.Absava
yav = yam - ebsiphon.ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebsiphon.tron_amo.Absamo
yam = ebsiphon.tron_amo.radamo + ebsiphon.tron_amo.conduit.Diametre
xav = ebsiphon.tron_amo.Absava
yav = ebsiphon.tron_amo.radava + ebsiphon.tron_amo.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin conduite aval
xam = ebsiphon.tron_amo.Absava + ls
yam = ebsiphon.tron_ava.radamo - ebsiphon.ds
xav = ebsiphon.tron_amo.Absava + ls + ebsiphon.ds
yav = ebsiphon.tron_ava.radamo
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2

xam = ebsiphon.tron_amo.Absava + ls + ebsiphon.ds
yam = ebsiphon.tron_ava.radamo
xav = xam + ebsiphon.tron_ava.conduit.Longueur
yav = ebsiphon.tron_ava.radava
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebsiphon.tron_amo.Absava + ls
yam = ebsiphon.tron_ava.radamo + ebsiphon.tron_ava.conduit.Diametre
xav = xam + ebsiphon.tron_ava.conduit.Longueur
yav = ebsiphon.tron_ava.radava + ebsiphon.tron_ava.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin du siphon ligne inférieure
xam = ebsiphon.tron_amo.Absava
yam = ebsiphon.tron_amo.radava - ebsiphon.ds
xav = ebsiphon.tron_amo.Absava + decalx
'yav = ebsiphon.tron_ava.radamo - ebsiphon.Ds
'yav = ebsiphon.tron_ava.radamo - decaly - ebsiphon.Ds
yav = decaly - ebsiphon.ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = xav + decalx
'yav = ebsiphon.tron_ava.radamo - decaly - ebsiphon.Ds
yav = decaly - ebsiphon.ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = xav + decalx
yav = yam
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = ebsiphon.tron_amo.Absava + ls
yav = ebsiphon.tron_ava.radamo - ebsiphon.ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin du siphon ligne supérieure
xam = ebsiphon.tron_amo.Absava
yam = ebsiphon.tron_amo.radava ' + ebsiphon.Ds
xav = ebsiphon.tron_amo.Absava
yav = ebsiphon.tron_amo.radava + ebsiphon.tron_amo.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebsiphon.tron_amo.Absava
yam = ebsiphon.tron_amo.radava  '+ ebsiphon.Ds
xav = ebsiphon.tron_amo.Absava + decalx
'yav = ebsiphon.tron_ava.radamo - decaly ' + ebsiphon.Ds
yav = decaly ' + ebsiphon.Ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = xav + decalx
'yav = (ebsiphon.tron_ava.radamo - decaly) '+ ebsiphon.Ds
yav = decaly '+ ebsiphon.Ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = xav + decalx
yav = yam
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xav
yam = yav
xav = ebsiphon.tron_amo.Absava + ls
yav = ebsiphon.tron_ava.radamo '+ ebsiphon.Ds
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebsiphon.tron_amo.Absava + ls
yam = yav
xav = ebsiphon.tron_amo.Absava + ls
yav = ebsiphon.tron_ava.radamo + ebsiphon.tron_ava.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
uc_g.redef_drwidth 1
'dessin des hauteurs d'eau
xam = ebsiphon.tron_amo.Absamo
yam = ebsiphon.tron_amo.radamo + h1
xav = ebsiphon.tron_amo.Absava
yav = ebsiphon.tron_amo.radava + h1
uc_g.dess_lign xam, yam, xav, yav, couleur.rouge, 1
xam = ebsiphon.tron_amo.Absava + ls
yam = ebsiphon.tron_ava.radamo + h2
xav = xam + ebsiphon.tron_ava.conduit.Longueur
yav = ebsiphon.tron_ava.radava + h2
uc_g.dess_lign xam, yam, xav, yav, couleur.rouge, 1

'dessin des lignes de pertes
xam = ebsiphon.tron_amo.Absava
yam = ebsiphon.tron_amo.radava + h1
xav = ebsiphon.tron_amo.Absava + ls / 2
yav = ebsiphon.tron_amo.radava + h1
uc_g.dess_lign_point xam, yam, xav, yav, couleur.rouge
xam = ebsiphon.tron_amo.Absava + ls / 2
yam = ebsiphon.tron_ava.radamo + h2
xav = ebsiphon.tron_amo.Absava + ls
yav = ebsiphon.tron_ava.radamo + h2
uc_g.dess_lign_point xam, yam, xav, yav, couleur.vert
xam = ebsiphon.tron_amo.Absava + ls / 2
yam = ebsiphon.tron_ava.radamo + h2 + jtot
xav = ebsiphon.tron_amo.Absava + ls
yav = ebsiphon.tron_ava.radamo + h2 + jtot
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
xam = ebsiphon.tron_amo.Absava + ls / 2
If ebsiphon.tron_amo.radava + h1 > ebsiphon.tron_ava.radamo + h2 + jtot Then
    yam = ebsiphon.tron_amo.radava + h1
Else
    yam = ebsiphon.tron_ava.radamo + h2 + jtot
End If
xav = ebsiphon.tron_amo.Absava + ls / 2
yav = ebsiphon.tron_ava.radamo + h2
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
'Cotation
uc_g.redef_drwidth 1
uc_g.dess_coth_text ebsiphon.tron_amo.Absava, ebsiphon.tron_amo.radava - ebsiphon.ds, _
ebsiphon.tron_amo.Absava + ls, ebsiphon.tron_ava.radamo - ebsiphon.ds, "Développé = " + ajout_zero(Trim(str(Round(ebsiphon.ls, 2)))) + " m", couleur_noir
uc_g.dess_text ebsiphon.tron_amo.Absava, _
ebsiphon.tron_amo.Absava + ls / 2, "-", ebsiphon.tron_amo.radava + ebsiphon.tron_amo.conduit.Diametre, "Perte de charge admissible " + ajout_zero(Trim(str(Round(ebsiphon.Jadm, 3)))) + " m", couleur.rouge
uc_g.dess_text ebsiphon.tron_amo.Absava + ls / 2, _
ebsiphon.tron_amo.Absava + ls, "+", ebsiphon.tron_ava.radamo + h2 + jtot, "Perte de charge totale " + ajout_zero(Trim(str(Round(jtot, 3)))) + " m", couleur.bleu
End Sub
Private Sub ini_lbresu()
'    Me.tb_resu.BackColor = &H8000000B
    Me.tb_resu.BorderStyle = 1
    Me.tb_resu.Text = ""
End Sub
Private Sub modi_lbresu()
'    Me.tb_resu.BackColor = &H80000009
    Me.tb_resu.BorderStyle = 1
End Sub

Private Sub init_graph(ByRef uc_graph As UC_graphique, ByVal h2 As Double, ByVal jtot As Double)
Dim ok As Boolean
Dim ecx As Double
Dim i As Integer
ok = False
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
    uc_graph.init_title
    uc_graph.init_titleh ""
    uc_graph.init_titleb ""
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX 0
uc_graph.init_MaxX ebsiphon.tron_amo.conduit.Longueur + ebsiphon.ls * 0.75 + ebsiphon.tron_ava.conduit.Longueur
uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
If ebsiphon.tron_ava.radamo + h2 + jtot <= ebsiphon.tron_amo.radamo + ebsiphon.tron_amo.conduit.Diametre Then
uc_graph.init_MaxY ebsiphon.tron_amo.radamo + ebsiphon.tron_amo.conduit.Diametre + 1#
Else
uc_graph.init_MaxY ebsiphon.tron_ava.radamo + h2 + jtot + 1#
End If
uc_graph.init_MinY Int(ebsiphon.tron_ava.radava) - 1.5
uc_graph.init_EchYn 1
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ouv_sauv = False
    Unload Frm_desprint
    Unload owner.fdessin
    owner.recharge_commentaire
End Sub

Private Sub Frm_Amont_Click()
Dim mes As String
Dim nom As String
nom = "Frm_Amont"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'owner.affich_aide Me.Name, "Siphon Conduite Amont"
End Sub
Private Sub Frm_Aval_Click()
Dim mes As String
Dim nom As String
nom = "Frm_Aval"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'owner.affich_aide Me.Name, "Siphon Conduite Aval"
End Sub

Private Sub Frm_siphon_Click()
Dim mes As String
Dim nom As String
nom = "Frm_siphon"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'    owner.affich_aide Me.Name, "Siphon Débit"

End Sub


Private Sub Lb_intsiph_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intsiph"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
'    owner.affich_aide Me.Name, "Siphon Débit"

End Sub
Private Sub Lb_intamo_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intamo"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
'owner.affich_aide Me.Name, "Siphon Conduite Amont"
End Sub
Private Sub Lb_intava_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intava"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
'owner.affich_aide Me.Name, "Siphon Conduite Aval"
End Sub

Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then
        Me.mnusave.Enabled = True
        Me.mnusaves.Enabled = True
        Me.mnusuppr.Enabled = True
'        Me.Mnuprint.Enabled = True
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
    reponse = MsgBox("Le siphon n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du siphon")
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
    reponse = MsgBox("Le siphon n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un siphon")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_siphon.Visible = True
            frmf.Label1.Caption = "Recherche d'un siphon "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_siphon_click
            End If
        Case Is = 7
'            Cb_siphon.Visible = True
            frmf.Label1.Caption = "Recherche d'un siphon "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_siphon_click
            End If
    End Select
Else
'    Cb_siphon.Visible = True
    frmf.Label1.Caption = "Recherche d'un siphon "
    frmf.Caption = nom
    frmf.Show 1
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_siphon_click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer, nb As Integer, j As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End

FrmPrint.Type1 = "siphon"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE SIPHON"
FrmPrint.sstitre1 = "Paramètres"
FrmPrint.ssTitre2 = "Données siphon"
FrmPrint.ssTitre4 = "Vérifications"
Frm_imp.Type1 = "siphon"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE SIPHON"
Frm_imp.sstitre1 = "Paramètres"
Frm_imp.ssTitre2 = "Données siphon"
Frm_imp.ssTitre4 = "Vérifications"
nb = (Tb_amo.count - 1) + 1
ReDim list_don1(nb, 5)
    list_don1(0, 1) = ""
    list_don1(0, 2) = Frm_Amont.Caption
    list_don1(0, 3) = ""
    list_don1(0, 4) = Frm_Aval.Caption
    list_don1(0, 5) = ""
For i = 0 To Tb_amo.count - 1
    list_don1(i + 1, 1) = Lb_intamo(i).Caption
    list_don1(i + 1, 2) = Tb_amo(i).Text
    list_don1(i + 1, 3) = Lb_uamo(i).Caption
    list_don1(i + 1, 4) = Tb_ava(i).Text
    list_don1(i + 1, 5) = Lb_uava(i).Caption
Next
nb = (Tb_siph.count - 1)
ReDim list_don2(nb, 3)
For i = 0 To nb
    list_don2(i, 1) = Lb_intsiph(i).Caption
    list_don2(i, 2) = Tb_siph(i).Text
    list_don2(i, 3) = Lb_usiph(i).Caption
Next
If Trim(Listcoud.coude(0).type) <> "" Then
    j = 0
    For i = 0 To 9
        If Trim(Listcoud.coude(i).type) <> "" Then
            j = j + 1
        End If
    Next
    ReDim list_don3(j, 5)
    j = 0
    list_don3(j, 1) = "Type"
    list_don3(j, 2) = "Nombre"
    list_don3(j, 3) = "Angle(d)"
    list_don3(j, 4) = "Rayon(m)"
    list_don3(j, 5) = ""
    For i = 0 To 9
        If Trim(Listcoud.coude(i).type) <> "" Then
            j = j + 1
            If Trim(Listcoud.coude(i).type) = "Arrondi" Then
                list_don3(j, 1) = "Coude arrondi"
                list_don3(j, 4) = Listcoud.coude(i).Rayon
            Else
                list_don3(j, 1) = "Coude à angle vif"
                list_don3(j, 4) = ""
            End If
            list_don3(j, 2) = Listcoud.coude(i).Nbre
            list_don3(j, 3) = Listcoud.coude(i).angle
            list_don3(j, 5) = ""
        End If
    Next
    FrmPrint.ssTitre3 = "Détail des singularités"
    Frm_imp.ssTitre3 = "Détail des singularités"
Else
    FrmPrint.ssTitre3 = ""
    Frm_imp.ssTitre3 = ""
End If
list_int1 = rec_list(tb_resu.Text)
'For i = 0 To UBound(list_int1)
'Debug.Print list_int1(i, 1), list_int1(i, 2), list_int1(i, 3)
'Next
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
'FrmPrint.Show
End Sub
Public Function lect_list(ByVal nom As String) As Variant
'For i = 0 To UBound(list_int1)
'Debug.Print list_int1(i, 1), list_int1(i, 2), list_int1(i, 3)
'Next
Select Case nom
Case Is = "list_don1"
    lect_list = list_don1
Case Is = "list_int1"
    lect_list = list_int1
Case Is = "list_don2"
    lect_list = list_don2
Case Is = "list_don3"
    lect_list = list_don3
End Select
End Function

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Public Sub reini_valeurs()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Call ini_lbresu
     If ebsiphon.dam > 0 And ebsiphon.iRadam > 0 And ebsiphon.Kam > 0 _
        And ebsiphon.Rdav > 0 Then
        Me.Cmd_amo.Enabled = True
    Else
        Me.Cmd_amo.Enabled = False
    End If
     If ebsiphon.dav > 0 And ebsiphon.iradav > 0 And ebsiphon.kav > 0 _
        And ebsiphon.Rdam > 0 Then
        Me.Cmd_ava.Enabled = True
    Else
        Me.Cmd_ava.Enabled = False
    End If
    If ebsiphon.dam > 0 And ebsiphon.iRadam > 0 And ebsiphon.Kam > 0 _
        And ebsiphon.dav > 0 And ebsiphon.iradav > 0 And ebsiphon.kav > 0 _
        And ebsiphon.Rdav > 0 And ebsiphon.Rdam > 0 And ebsiphon.ds > 0 _
        And ebsiphon.Ks > 0 And ebsiphon.Qmax > 0 And ebsiphon.ls > 0 _
        And ebsiphon.Kc > 0 Then
        Me.Cmd_calcul.Enabled = True
        ' impression false
                    Me.mnuprint.Enabled = False
    Else
        Me.Cmd_calcul.Enabled = False
        ' impression false
                    Me.mnuprint.Enabled = False
    End If
    ouv_sauve = True
End Sub
Public Sub ini_ebsiphon()
    ebsiphon.dam = 0
    ebsiphon.iRadam = 0
    ebsiphon.Kam = 0
    ebsiphon.dav = 0
    ebsiphon.kav = 0
    ebsiphon.iradav = 0
    ebsiphon.Rdav = 0#
    ebsiphon.Rdam = 0#
    ebsiphon.Jadm = 0#
    ebsiphon.ds = 0#
    ebsiphon.Ks = 0
    ebsiphon.Qmax = 0#
    ebsiphon.ls = 0#
    ebsiphon.Kc = 0
    ebsiphon.List_coude = Listcoud
    ebsiphon.Ipl = 0#
    ebsiphon.deltaH1 = 0#
    ebsiphon.deltaH2 = 0#
    ebsiphon.IPs = 0#
    ebsiphon.tron_amo.Absamo = 0
    ebsiphon.tron_amo.Absava = 0
    ebsiphon.tron_amo.radamo = 0#
    ebsiphon.tron_amo.radava = 0#
    ebsiphon.tron_amo.conduit.Diametre = 0
    ebsiphon.tron_amo.conduit.Longueur = 0
    ebsiphon.tron_amo.conduit.pente = 0
    ebsiphon.tron_amo.conduit.rugosite = 0
    ebsiphon.tron_amo.conduit.typ = 0
    ebsiphon.tron_ava.Absamo = 0
    ebsiphon.tron_ava.Absava = 0
    ebsiphon.tron_ava.radamo = 0#
    ebsiphon.tron_ava.radava = 0#
    ebsiphon.tron_ava.conduit.Diametre = 0
    ebsiphon.tron_ava.conduit.Longueur = 0
    ebsiphon.tron_ava.conduit.pente = 0
    ebsiphon.tron_ava.conduit.rugosite = 0
    ebsiphon.tron_ava.conduit.typ = 0
End Sub
Private Sub ini_form()
    Me.Tb_amo(0).Text = rempl_virgule(Format(ebsiphon.dam, "###0"))
    Me.Tb_amo(1).Text = rempl_virgule(Format(ebsiphon.iRadam, "###0"))
    Me.Tb_amo(2).Text = rempl_virgule(Format(ebsiphon.Kam, "###0"))
    Me.Tb_ava(0).Text = rempl_virgule(Format(ebsiphon.dav, "###0"))
    Me.Tb_ava(1).Text = rempl_virgule(Format(ebsiphon.iradav, "###0"))
    Me.Tb_ava(2).Text = rempl_virgule(Format(ebsiphon.kav, "###0"))
    Me.Tb_amo(3).Text = rempl_virgule(Format(ebsiphon.Rdav, "##0.00"))
    Me.Tb_ava(3).Text = rempl_virgule(Format(ebsiphon.Rdam, "##0.00"))
    Me.Tb_siph(4).Text = rempl_virgule(Format(ebsiphon.Kc, "#0.000"))
    Me.Tb_siph(0).Text = rempl_virgule(Format(ebsiphon.ds * 1000, "###0"))
    Me.Tb_siph(1).Text = rempl_virgule(Format(ebsiphon.Ks, "###0"))
    Me.Tb_siph(2).Text = rempl_virgule(Format(ebsiphon.Qmax, "#0.000"))
    Me.Tb_siph(3).Text = rempl_virgule(Format(ebsiphon.ls, "####0.00"))
End Sub

Public Sub ini_listcoud()
Dim coude As st_Coude
coude.type = ""
coude.Nbre = 0
coude.angle = 0#
coude.Rayon = 0#
For i = 1 To 10
    Listcoud.coude(i - 1) = coude
Next
End Sub
Private Sub Tb_amo_Change(Index As Integer)
 Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie diamètre conduite Amont", "I")
            Case Is = 1
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie pente conduite Amont", "I")
            Case Is = 2
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie coefficient conduite Amont", "I")
            Case Is = 3
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie cote radier Aval", "R")
        End Select
  If nom = "" Then
    Tb_amo(Index).Text = sval_champ
    Tb_amo(Index).SelStart = iSels
    Tb_amo(Index).SelLength = iSell
  End If
End If
'****

   Select Case Index
        Case Is = 0
            ebsiphon.dam = txtVersNum(Me.Tb_amo(0).Text)
        Case Is = 1
            ebsiphon.iRadam = txtVersNum(Me.Tb_amo(1).Text)
        Case Is = 2
            ebsiphon.Kam = txtVersNum(Me.Tb_amo(2).Text)
        Case Is = 3
            ebsiphon.Rdav = txtVersNum(Me.Tb_amo(3).Text)
    End Select
    Call reini_valeurs
     sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_amo_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_amo"
Call sel_text(Tb_amo(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_amo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_amo(Index).Text
    iSels = Tb_amo(Index).SelStart
    iSell = Tb_amo(Index).SelLength
    bKP = True
'   If Len(Tb_amo(Index).Text) <= Tb_amo(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie diamètre conduite Amont", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie pente conduite Amont", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie coefficient conduite Amont", "I")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie cote radier Aval", "R")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_amo_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_amo"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
DoEvents
Me.Show
Call sel_text(Tb_amo(Index))
'owner.affich_aide Me.Name, "Siphon Conduite Amont"
End Sub

Private Sub Tb_amo_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_amo", Index, txtVersNum(Tb_amo(Index).Text))
    If Not ok Then
        Tb_amo(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_ava_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_ava(Index).Text, "Saisie  diamètre conduite Aval", "I")
            Case Is = 1
                nom = verif_cart0(Tb_ava(Index).Text, "Saisie pente conduite Aval", "I")
            Case Is = 2
                nom = verif_cart0(Tb_ava(Index).Text, "Saisie coefficient conduite Aval", "I")
            Case Is = 3
                nom = verif_cart0(Tb_ava(Index).Text, "Saisie cote radier Amont", "R")
        End Select
  If nom = "" Then
    Tb_ava(Index).Text = sval_champ
    Tb_ava(Index).SelStart = iSels
    Tb_ava(Index).SelLength = iSell
  End If
End If
'****

    Select Case Index
        Case Is = 0
            ebsiphon.dav = txtVersNum(Me.Tb_ava(0).Text)
        Case Is = 1
            ebsiphon.iradav = txtVersNum(Me.Tb_ava(1).Text)
        Case Is = 2
            ebsiphon.kav = txtVersNum(Me.Tb_ava(2).Text)
        Case Is = 3
            ebsiphon.Rdam = txtVersNum(Me.Tb_ava(3).Text)
    End Select
    Call reini_valeurs
     sval_champ = ""
    bKP = False

End Sub
Private Sub Tb_ava_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_ava"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
DoEvents
Me.Show
Call sel_text(Tb_ava(Index))
'owner.affich_aide Me.Name, "Siphon Conduite Aval"
End Sub

Private Sub Tb_ava_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_ava"
Call sel_text(Tb_ava(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_ava_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_ava(Index).Text
    iSels = Tb_ava(Index).SelStart
    iSell = Tb_ava(Index).SelLength
    bKP = True
'    If Len(Tb_ava(Index).Text) <= Tb_ava(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie  diamètre conduite Aval", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie pente conduite Aval", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie coefficient conduite Aval", "I")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie cote radier Amont", "R")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_ava_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_ava", Index, txtVersNum(Tb_ava(Index).Text))
    If Not ok Then
        Tb_ava(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_siph_change(Index As Integer)
Dim nom As String

If bKP Then
    Select Case Index
        Case Is = 0
            nom = verif_cart0(Tb_siph(Index).Text, "Saisie diamètre du siphon", "R")
        Case Is = 1
            nom = verif_cart0(Tb_siph(Index).Text, "Saisie coefficient siphon", "I")
        Case Is = 2
            nom = verif_cart0(Tb_siph(Index).Text, "Saisie débit maxi du siphon", "R")
        Case Is = 3
            nom = verif_cart0(Tb_siph(Index).Text, "Saisie longueur développée du siphon", "R")
        Case Is = 4
            nom = verif_cart0(Tb_siph(Index).Text, "Saisie coefficient singularités", "R")
    End Select
  If nom = "" Then
    Tb_siph(Index).Text = sval_champ
    Tb_siph(Index).SelStart = iSels
    Tb_siph(Index).SelLength = iSell
  End If
End If
'****

    Select Case Index
        Case Is = 0
            ebsiphon.ds = txtVersNum(Me.Tb_siph(0).Text) / 1000#
            If ebsiphon.ds > 0 Then
                Call calc_kc
            End If
        Case Is = 1
            ebsiphon.Ks = txtVersNum(Me.Tb_siph(1).Text)
        Case Is = 2
            ebsiphon.Qmax = txtVersNum(Me.Tb_siph(2).Text)
        Case Is = 3
            ebsiphon.ls = txtVersNum(Me.Tb_siph(3).Text)
        Case Is = 4
            ebsiphon.Kc = txtVersNum(Me.Tb_siph(4).Text)
    End Select
    Call reini_valeurs
    sval_champ = ""
    bKP = False

End Sub
Private Sub Tb_siph_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_siph"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
DoEvents
Me.Show
Call sel_text(Tb_siph(Index))
'    owner.affich_aide Me.Name, "Siphon Débit"
End Sub

Private Sub Tb_siph_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_siph"
Call sel_text(Tb_siph(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_siph_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_siph(Index).Text
    iSels = Tb_siph(Index).SelStart
    iSell = Tb_siph(Index).SelLength
    bKP = True
'    If Len(Tb_siph(Index).Text) <= Tb_siph(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'        KeyAscii = verif_car(Tb_siph(Index).Text, KeyAscii, "Saisie diamètre du siphon", "R")
'            Case Is = 1
'        KeyAscii = verif_car(Tb_siph(Index).Text, KeyAscii, "Saisie coefficient siphon", "I")
'            Case Is = 2
'        KeyAscii = verif_car(Tb_siph(Index).Text, KeyAscii, "Saisie débit maxi du siphon", "R")
'            Case Is = 3
'        KeyAscii = verif_car(Tb_siph(Index).Text, KeyAscii, "Saisie longueur développée du siphon", "R")
'            Case Is = 4
'        KeyAscii = verif_car(Tb_siph(Index).Text, KeyAscii, "Saisie coefficient singularités", "R")
'        End Select
'    End If
End If
End Sub

Private Sub calcul_amont_aval()
Dim z1 As Double, z2 As Double, h1 As Double, h2 As Double, sresult As String
Dim Jadm As Double, Ipl As Double, deltaH1 As Double, deltaH2 As Double, IPs As Double
Dim v0 As Double, v1 As Double, v2 As Double, Q As Double, surf As Double
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
Dim res_amo As debit_conduit
Dim res_ava As debit_conduit
Dim cana_ava As conduite
' conduite amont -> troncon amont
Dim message As String
message = ""
If ebsiphon.Kam < 50 Or ebsiphon.Kam > 80 Then
    message = message + "Coefficient de rugosité hors plage 50 - 80 (conduite amont"
End If
If ebsiphon.kav < 50 Or ebsiphon.kav > 80 Then
    If message = "" Then
        message = message + "Coefficient de rugosité hors plage 50 - 80 (conduite aval"
    Else
        message = message + ", conduite aval"
    End If
End If
If ebsiphon.Ks < 50 Or ebsiphon.Ks > 80 Then
    If message = "" Then
        message = message + "Coefficient de rugosité hors plage 50 - 80 (siphon"
    Else
        message = message + ", siphon"
    End If
End If
If message <> "" Then
    message = message + ")"
    MsgBox message, vbExclamation, "Calcul SIPHON"
End If

    cana_amo.Diametre = ebsiphon.dam / 1000#
    cana_amo.Longueur = 5
    cana_amo.pente = ebsiphon.iRadam / 10000#
    cana_amo.rugosite = ebsiphon.Kam
    cana_amo.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + cana_amo.Longueur
      .conduit = cana_amo
      .radava = ebsiphon.Rdav
      .radamo = ebsiphon.Rdav + 0.3 ' cana_amo.Longueur * cana_amo.pente
    End With
    ebsiphon.tron_amo = troamo
    res_amo = calc_debit_tr(ebsiphon.tron_amo, ebsiphon.Qmax)
    cana_ava.Diametre = ebsiphon.dav / 1000#
    cana_ava.Longueur = 5
    cana_ava.pente = ebsiphon.iradav / 10000#
    cana_ava.rugosite = ebsiphon.kav
    cana_ava.typ = 2
    With troava
      .Absava = 0#
      .Absava = .Absava + cana_ava.Longueur
      .conduit = cana_ava
      .radamo = ebsiphon.Rdam
      .radava = ebsiphon.Rdam - 0.3 'cana_ava.Longueur * cana_ava.pente
    End With
    ebsiphon.tron_ava = troava
    res_ava = calc_debit_tr(ebsiphon.tron_ava, ebsiphon.Qmax)
    h1 = res_amo.hauteur
    h2 = res_ava.hauteur
    z1 = ebsiphon.tron_amo.radava + h1
    z2 = ebsiphon.tron_ava.radamo + h2
    v1 = res_amo.vitesse
    v2 = res_ava.vitesse
    Jadm = z1 - z2
    ebsiphon.Jadm = Round(Jadm, 3)
    'a revoir perte de charge linéaire
    Dim lo As conduite
    lo.Diametre = ebsiphon.ds
    lo.rugosite = ebsiphon.Ks
    lo.typ = 2
    p = pent_mot0(lo, ebsiphon.Qmax)
    Ipl = 10.2938 * ebsiphon.ls * ((ebsiphon.Qmax / (ebsiphon.Ks * (ebsiphon.ds ^ (8# / 3#)))) ^ 2)
    Q = ebsiphon.Qmax
    surf = pi * (ebsiphon.ds / 2) ^ 2
    v0 = Q / surf
' houpie 20040518
    deltaH1 = 0.5 * ((v0 ^ 2) / (2 * 9.81))
' certu 200480828
'    deltaH2 = (v2 ^ 2 - v0 ^ 2) / (2 * 9.81)
    deltaH2 = (v0 - v2) ^ 2 / (2 * 9.81)
    ebsiphon.Ipl = Round(Ipl, 3)
    
    ebsiphon.deltaH1 = Round(deltaH1, 3)
    ebsiphon.deltaH2 = Round(deltaH2, 3)
 ' definition des singularites Kc = somme (Ksi)
  'tableau type nombre angle rayon (rayon = 0 si coude)
    IPs = ebsiphon.Kc * v0 ^ 2 / (2 * 9.81)
    ebsiphon.IPs = Round(IPs, 3)
    Jtotal = deltaH1 + deltaH2 + Ipl + IPs
    Call modi_lbresu
    sresult = "  Vitesse dans le siphon                 = " + ajout_zero(Trim(str(Round(v0, 2)))) + " m/s"
    
    sresult = sresult + Chr(13) + Chr(10) + "  Perte de charge linéaire               = " + ajout_zero(Trim(str(Round(Ipl, 3)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  Perte de charge en entrée              = " + ajout_zero(Trim(str(Round(deltaH1, 3)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  Perte de charge en sortie              = " + ajout_zero(Trim(str(Round(deltaH2, 3)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  Perte de charge dans les coudes        = " + ajout_zero(Trim(str(Round(IPs, 3)))) + " m"
'    sresult = sresult + Chr(13) + Chr(10)
    sresult = sresult + Chr(13) + Chr(10) + "  - Perte de charge admissible             = " + ajout_zero(Trim(str(Round(Jadm, 3)))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "  - Perte de charge totale                 = " + ajout_zero(Trim(str(Round(Jtotal, 3)))) + " m"
    Me.tb_resu.Text = sresult
    Me.Cmd_calcul.Enabled = False
        ' impression true
                    Me.mnuprint.Enabled = True
    Call init_graph(owner.fdessin.UC_graphique1, h2, Jtotal)
    Call dess_siphon(owner.fdessin.UC_graphique1, h1, h2, Jtotal)
    Call init_graph(Frm_desprint.UC_graphique1, h2, Jtotal)
    Call dess_siphon(Frm_desprint.UC_graphique1, h1, h2, Jtotal)
 
End Sub
Public Sub calc_kc()
Dim i As Integer, k As Double, Kc As Double, alpha As Double, ray As Double, rap As Double
Dim alpha1 As Double
Kc = 0#
For i = 1 To 10
    If Trim(Listcoud.coude(i - 1).type) = "Arrondi" Then
        If Listcoud.coude(i - 1).Rayon > 0 Then
            alpha = Listcoud.coude(i - 1).angle / 90
            rap = ebsiphon.ds / (2 * Listcoud.coude(i - 1).Rayon)
            k = alpha * (0.131 + 1.847 * (rap ^ 3.5))
            Kc = Kc + k * Listcoud.coude(i - 1).Nbre
        Else
            alpha = Listcoud.coude(i - 1).angle
            alpha1 = pi * alpha / 180#
            k = 0.946 * Sin((alpha1 / 2)) ^ 2 + 2.05 * Sin((alpha1 / 2)) ^ 4
            Kc = Kc + k * Listcoud.coude(i - 1).Nbre
        End If
    End If
    If Trim(Listcoud.coude(i - 1).type) = "Angle vif" Then
        alpha = Listcoud.coude(i - 1).angle
        alpha1 = pi * alpha / 180#
        k = 0.946 * Sin((alpha1 / 2)) ^ 2 + 2.05 * Sin((alpha1 / 2)) ^ 4
        Kc = Kc + k * Listcoud.coude(i - 1).Nbre
    End If
Next
    Me.Tb_siph(4).Text = rempl_virgule(Format(Round(Kc, 3), "####0.000"))
    ebsiphon.Kc = Round(Kc, 3)
    
End Sub
Private Sub mnusaves_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
'Debug.Print fich_lect_edit
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'un siphon "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du siphon " & Me.Tb_titre.Text
'         Frm_titre.Label3.Caption = " de l'étude " & fich_lect
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du siphon (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_savsi
Dim za1 As st_savsi1
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
            za = za1.stsavsi
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'une siphon")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "siphon"
            za.nom = Tb_titre.Text
            za.siphon = ebsiphon
            za1.stsavsi = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        za.type = "siphon"
        za.nom = Tb_titre.Text
        za.siphon = ebsiphon
        za1.stsavsi = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        si_texte = Trim(Tb_titre.Text)
        Cb_siphon.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom du siphon n'est pas renseigné.", , "Sauvegarde d'un siphon")
End If
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If

Call flockb(nom_fich)
End Sub
Private Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If
End Sub
Private Sub mnusuppr_Click()
Dim za As st_savsi
Dim za1 As st_savsi1
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_siphon.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_siphon.Text) + " va être supprimé .", 4, "Suppression d'un siphon")
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
            za = za1.stsavsi
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_siphon.Text)) Then
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
    Call ini_listcoud
    Call ini_ebsiphon
    Call ini_form
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le siphon n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du siphon")
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
Private Sub lect_fich()
Dim za As st_savsi
Dim za1 As st_savsi1
Call funlockb

    lhFicDbf = FreeFile
    Cb_siphon.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavsi
        If Trim(za.type) = nom_type Then
            Cb_siphon.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
si_texte = Cb_siphon.list(0)
Cb_siphon.Text = Cb_siphon.list(0)
Cb_siphon.Refresh
 
Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Sub Cb_siphon_click()
Dim za As st_savsi
Dim za1 As st_savsi1
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    si_texte = Trim(nom_ouvrage)
    Cb_siphon.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavsi
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_siphon.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebsiphon = za.siphon
            Listcoud = ebsiphon.List_coude
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
'Debug.Print fich_lect_edit
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
Public Sub Init_ss_commentaire()
   owner.affich_aide Me.Name, ""  'Dimensionnement d'un siphon"
End Sub


Private Sub Tb_siph_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_siph", Index, txtVersNum(Tb_siph(Index).Text))
    If Not ok Then
        Tb_siph(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub
