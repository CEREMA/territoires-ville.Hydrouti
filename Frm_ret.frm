VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_ret 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Bassin de Rétention"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_ret.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Méthode des pluies"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du bassin de rétention méthode des pluies"
      Top             =   50
      Width           =   2055
   End
   Begin VB.TextBox Tb_volume 
      BackColor       =   &H80000016&
      Height          =   405
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1950
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox tb_resu 
      Height          =   1575
      Left            =   5040
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483626
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frm_ret.frx":08CA
   End
   Begin VB.Frame Frm_rect 
      Height          =   1695
      Left            =   5040
      TabIndex        =   23
      Top             =   2280
      Width           =   4455
      Begin VB.CommandButton Cmd_graph 
         Caption         =   "Graphique"
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   960
         Width           =   1000
      End
      Begin VB.TextBox Tb_long 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   29
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox Tb_larg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   30
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Tb_prof 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   31
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox Tb_rap 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   32
         Top             =   1320
         Width           =   900
      End
      Begin VB.CheckBox Chk_long 
         Height          =   200
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   290
         Width           =   200
      End
      Begin VB.CheckBox Chk_larg 
         Height          =   200
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   650
         Width           =   200
      End
      Begin VB.CheckBox Chk_prof 
         Height          =   200
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1005
         Width           =   200
      End
      Begin VB.CheckBox Chk_rap 
         Height          =   200
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1370
         Width           =   200
      End
      Begin VB.CommandButton Cmd_schema 
         Caption         =   "Schéma"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Dimensionnement et dessin du bassin de rétention"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.Label Lb_intlong 
         Caption         =   "Longueur"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Lb_intlarg 
         Caption         =   "Largeur"
         Height          =   255
         Left            =   580
         TabIndex        =   39
         Top             =   650
         Width           =   1215
      End
      Begin VB.Label Lb_intprof 
         Caption         =   "Hauteur d'eau"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label Lb_intrap 
         Caption         =   "Rapport l/h"
         Height          =   255
         Left            =   580
         TabIndex        =   37
         Top             =   1370
         Width           =   1215
      End
      Begin VB.Label Lb_ulong 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   290
         Width           =   300
      End
      Begin VB.Label Lb_ularg 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Top             =   650
         Width           =   300
      End
      Begin VB.Label Lb_uprof 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   1010
         Width           =   300
      End
      Begin VB.Label Lb_urap 
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   1365
         Width           =   105
      End
   End
   Begin VB.CommandButton Cmd_hydro 
      Caption         =   "Hydrogramme"
      Height          =   255
      Left            =   7560
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du bassin de rétention à partir de l'hydrogramme du bassin versant  sélectionné "
      Top             =   50
      Width           =   1935
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   5400
      MaxLength       =   30
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox Cb_retention 
      Height          =   315
      Left            =   360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   4000
   End
   Begin VB.Frame Frm_parm 
      Caption         =   "Paramètres pluviométriques pour un résultat en mm/mn"
      Height          =   1215
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   4335
      Begin VB.TextBox Tb_par 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   3550
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "seuil"
         Top             =   720
         Width           =   645
      End
      Begin VB.TextBox Tb_par 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2750
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "b2"
         Top             =   720
         Width           =   645
      End
      Begin VB.TextBox Tb_par 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "a2"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox Tb_par 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "a1"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox Tb_par 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2050
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "b1"
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "b"
         Height          =   255
         Left            =   1800
         TabIndex        =   48
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "a"
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Lb_intpar 
         Caption         =   "Seuil (mn)"
         Height          =   225
         Index           =   2
         Left            =   3480
         TabIndex        =   46
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lb_upar 
         Height          =   135
         Index           =   1
         Left            =   3555
         TabIndex        =   44
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Lb_upar 
         Height          =   255
         Index           =   0
         Left            =   3555
         TabIndex        =   43
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lb_intpar 
         Caption         =   "Coefficients  Montana"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Lb_intpar 
         Height          =   45
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.TextBox Tb_Qf 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2780
      MaxLength       =   6
      TabIndex        =   8
      Top             =   3600
      Width           =   900
   End
   Begin VB.CommandButton Cmd_Sel_Bv 
      Caption         =   " Sélection d'un bassin versant"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Frame Frm_bassin 
      Caption         =   "Bassin versant "
      Height          =   2175
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2445
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2445
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "%"
         Height          =   255
         Index           =   1
         Left            =   3555
         TabIndex        =   21
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "ha"
         Height          =   255
         Index           =   0
         Left            =   3555
         TabIndex        =   13
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Coefficient d'apport du B.V."
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   765
         Width           =   2055
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Surface du B.V."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   405
         Width           =   2055
      End
   End
   Begin VB.Label Lb_uqf 
      Caption         =   "l/s"
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   3645
      Width           =   375
   End
   Begin VB.Label Lb_Qf 
      Caption         =   "Débit de fuite de la retenue"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   3645
      Width           =   2055
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Bassin de rétention"
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
Attribute VB_Name = "Frm_ret"
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
Private chang_long As Boolean
Private chang_larg As Boolean
Private chang_prof As Boolean
Private chang_rap As Boolean
Private ebret_dess As ret_dess
Private nombassin As String
Private list_don1() As Variant
Private list_int1() As Variant
Private list_resu1() As Variant
Private type_calcul As String * 1
Private ret_texte As String
Private fen_titre As String
Public titre_sav As String
Private list_tb() As Variant
Private sval_champ As String
Private bKP As Boolean
Private iSels As Integer
Private iSell As Integer
Private label_prec As String
Private mes_prec As String
Private index_prec As Integer
Private change_coul As Boolean

Private bkprof As Boolean
Private bklarg  As Boolean
Private bklong  As Boolean
Private bkrap As Boolean


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
'    Case Is = "Tb_bv"
'         nom1 = "Lb_intbv"
'    Case Is = "Tb_par"
'         nom1 = "Lb_intpar"
'    Case Is = "Tb_Qf"
'         nom1 = "Lb_Qf"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intbv"
'         Lb_intbv(index_prec).ForeColor = coulp
'    Case Is = "Lb_intpar"
'         Lb_intpar(index_prec).ForeColor = coulp
'    Case Is = "Lb_Qf"
'         Lb_Qf.ForeColor = coulp
'    Case Is = "Frm_bassin"
'         Frm_bassin.ForeColor = coulp
'    Case Is = "Frm_parm"
'         Frm_parm.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intbv"
'         Lb_intbv(Index).ForeColor = coul
'    Case Is = "Lb_intpar"
'         Lb_intpar(Index).ForeColor = coul
'    Case Is = "Lb_Qf"
'         Lb_Qf.ForeColor = coul
'    Case Is = "Frm_bassin"
'         Frm_bassin.ForeColor = coul
'   Case Is = "Frm_parm"
'         Frm_parm.ForeColor = coul
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
    Case Is = "Lb_intbv"
         Tb_bv(Index).SetFocus
    Case Is = "Lb_intpar"
         Tb_par(Index).SetFocus
    Case Is = "Lb_Qf"
         Tb_Qf.SetFocus
    Case Is = "Frm_bassin"
         Tb_bv(0).SetFocus
    Case Is = "Frm_parm"
         Tb_par(0).SetFocus
    Case Is = "Cmd_calcul"
         Cmd_calcul.SetFocus
    Case Is = "Cmd_hydro"
         Cmd_hydro.SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_intbv", "Tb_bv", "Frm_bassin", "Lb_Qf", "Tb_Qf", "Cmd_calcul"
        mes = IDhlp_RetentionDimensionnementMethodePluies  '"Dimensionnement par la  méthode des pluies"
'         mes = "par la méthode des pluies"
    Case Is = "Lb_intpar", "Tb_par", "Frm_parm"
        mes = IDhlp_RetentionCoefficientsMontana  '"Choix des coefficients a et b de Montana"
    Case Is = "Cmd_hydro"
        mes = IDhlp_RetentionDimensionnementMethodeHydrogramme  '"Dimensionnement par la  méthode de l'hydrogramme"
         DoEvents
'         mes = "par la  méthode de l'hydrogramme"
End Select
mes_prec = mes
Rec_Mes = mes
End Function
Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant  ', l2() As Variant
l0 = Array(0)
l1 = Array(0, "TB_bv", "TB_par", "TB_qf", "CMD_calcul", "CMD_hydro", "Tb_long", "Tb_larg", "Tb_prof", "Tb_rap", "Cmd_graph")
'l2 = Array(0,"TB_par_ep", "TB_par_eu", "TB_par_pl")
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1)) ', 0 To UBound(l2))
list_tb = Array(l0, l1)  ', l2)

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
  ' Me.Width = 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub
Private Sub dess_stock()
Dim ebvolume As volume_dess
ebvolume.coef = ebret_dess.coef
ebvolume.Largeur = ebret_dess.Largeur
ebvolume.Longueur = ebret_dess.Longueur
ebvolume.Profondeur = ebret_dess.Profondeur
ebvolume.Rapport = ebret_dess.Rapport
Select Case ebret_dess.type
    Case Is = "rect"
 '      owner.fdessin.UC_graphique2.graphique_clear
        If ebret_dess.Longueur > 0 And ebret_dess.Largeur > 0 And ebret_dess.Profondeur > 0 Then
            Call init_graph_rect(owner.fdessin.UC_graphique2, ebvolume)
            Call dess_stock_rect(owner.fdessin.UC_graphique2, ebvolume)
            Call init_graph_rect(Frm_desprint.UC_graphique1, ebvolume)
            Call dess_stock_rect(Frm_desprint.UC_graphique1, ebvolume)
         ' impression true
                    Me.mnuprint.Enabled = True
            Cmd_schema.Enabled = True
            owner.fdessin.UC_graphique1.Visible = False
            owner.fdessin.UC_graphique2.Visible = True
        End If
'    Case Is = "circ"
'        owner.fdessin.UC_graphique2.graphique_clear
'        Call init_graph_circ(owner.fdessin.UC_graphique2, ebvolume)
'        Call dess_stock_circ(owner.fdessin.UC_graphique2, ebvolume)
'
End Select
    ouv_sauve = True
End Sub
Private Sub ini_form()
'Houpie 2005/03/21
'    Me.Tb_bv(0).Text = rempl_virgule(Format(ebret.surface, "###0"))
 '   Me.Tb_bv(0).Text = ajout_zero(Trim(Str(ebret.surface)))
    Me.Tb_bv(0).Text = rempl_virgule(Format(ebret.surface, "###0.00"))
    Me.Tb_bv(1).Text = rempl_virgule(Format(ebret.Ca, "###0"))
    Me.Tb_Qf.Text = rempl_virgule(Format(ebret.qf, "###0.00"))
    Me.Tb_par(0).Text = rempl_virgule(Format(ebret.amontana, "##0.000"))
    Me.Tb_par(1).Text = rempl_virgule(Format(ebret.bmontana, "##0.000"))
    Me.Tb_par(2).Text = rempl_virgule(Format(ebret.a1montana, "##0.000"))
    Me.Tb_par(3).Text = rempl_virgule(Format(ebret.b1montana, "##0.000"))
    Me.Tb_par(4).Text = rempl_virgule(Format(ebret.Seuil, "####0"))
    Call init_form_dess
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
End Sub
Private Sub Opt_rect()
    owner.fdessin.UC_graphique2.graphique_clear
    ebret_dess.type = "rect"
    ebret_dess.coef = 0.5 'difference entre hauteur et hauteur d'eau
    chang_long = Chk_long.Enabled
    chang_larg = Chk_larg.Enabled
    chang_prof = Chk_prof.Enabled
    chang_rap = Chk_rap.Enabled
    Me.Frm_rect.Visible = True
    Call dess_stock
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
Private Sub lect_fich()
Dim za As st_savret
Dim za1 As st_savret1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_retention.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavret
        If Trim(za.type) = nom_type Then
            Cb_retention.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
ret_texte = Cb_retention.list(0)
Cb_retention.Text = Cb_retention.list(0)
Cb_retention.Refresh
 
Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
Call flockb(nom_fich)
End Sub

Private Sub Cb_retention_Change()
    Cb_retention.Text = ret_texte
End Sub

Public Sub Cb_retention_click()
Dim za As st_savret
Dim za1 As st_savret1
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    ret_texte = Trim(nom_ouvrage)
    Cb_retention.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavret
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_retention.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebret = za.retention
            ebret_dess = ebret.desssret
            Me.Frm_rect.Visible = False

            Call ini_form
'           Call reini_valeurs
        Me.Frm_rect.Visible = False

            Call ini_tbresu
            nombassin = ebret.nombv
            Me.Frm_bassin.Caption = "Bassin versant : " + nombassin 'ebret.nom
            type_calcul = ebret.type_calcul
'        Me.Cmd_del.Visible = True
            If Trim(nombassin) <> "" Then
                 Close #lhFicDbf
                 If rec_bassin(nombassin, "versant") Then
                    Me.Cmd_hydro.Visible = True
                    Me.Cmd_calcul.Visible = True
                 Else
                    Me.Cmd_hydro.Visible = False
                    Me.Cmd_calcul.Visible = True
                End If
            Else
                 Me.Cmd_hydro.Visible = False
                 Me.Cmd_calcul.Visible = True
            End If
            If type_calcul = "V" Then
                If Cmd_calcul.Enabled Then
                    Call Calc_volume
                    Call reini_dess
                    owner.fdessin.UC_graphique1.Visible = True
                    owner.fdessin.UC_graphique2.Visible = False
                End If
            Else
            If Cmd_hydro.Enabled Then
                ehyd.qfuite = ebret.qf
                Call calc_hydro
                Call reini_dess
                owner.fdessin.UC_graphique1.Visible = True
                owner.fdessin.UC_graphique2.Visible = False
            End If
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

Private Sub Cb_retention_KeyDown(KeyCode As Integer, Shift As Integer)
    ret_texte = Cb_retention.Text
    Cb_retention.Text = ret_texte
End Sub

Private Sub Cb_retention_KeyPress(KeyAscii As Integer)
    ret_texte = Cb_retention.Text
End Sub

Private Sub Chk_larg_Click()
If Me.Chk_long.Enabled Then
    Call check_volume
End If
End Sub

Private Sub Cmd_graph_Click()
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
End Sub

Private Sub Cmd_hydro_Click()
    Dim mes As String
    Dim nom As String
    nom = "Cmd_hydro"
    Call key13(Me)
    type_calcul = "H"
    Call ini_tbresu
    ehyd.qfuite = ebret.qf
    Call calc_hydro
    Call reini_dess
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
    nom = "Cmd_hydro"
    mes = Rec_Mes(nom, 0)
    Change_Focus nom, 0
    owner.affich_aide Me.Name, mes
End Sub



Private Sub Cmd_schema_Click()
    owner.fdessin.UC_graphique1.Visible = False
    owner.fdessin.UC_graphique2.Visible = True
End Sub

Private Sub Cmd_Sel_Bv_Click()
    Dim pict1 As New StdPicture
    dess_anc = chemin_app + "dessanc.bmp"
    If Dir(dess_anc) <> "" Then
        Kill dess_anc
    End If
    Set pict1 = owner.fdessin.UC_graphique1.lire_pict1()
    SavePicture pict1, chemin_app + "dessanc.bmp"
    owner.fdessin.UC_graphique1.graphique_clear
    Me.Enabled = False
    ret_bv = True
    Set owner.fbassin = New Frm_bv2
    owner.fbassin.Show
    owner.fbassin.nom_ouvrage = nombassin
    owner.fbassin.Cmd_retour.Visible = True
    owner.fbassin.Cmd_retour.Caption = "Retour au bassin de rétention"
    '20040324
    fich_lect = nom_fich
    Call owner.fbassin.rec_bassin_versant
    owner.affich_aide owner.fbassin.Name, "Module" '"Calcul de débit de bassin versant"
End Sub

Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
'Me.Tb_par(1).ToolTipText = ""
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'un bassin de rétention"
    Change_Couleur "Me", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Form_Load()
    okg = True
    Me.KeyPreview = True
    Call ini_tooltip_ret(Me)
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'''''    owner.affich_aide Me.Name, "Rétention"
'    nom_fich = chemin_app + "retention.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "retention"
    fen_titre = Me.Caption
    ouv_sauve = False
    save_fich = False
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Frm_desprint.Show
    Cb_retention.Visible = False
    Frm_desprint.Visible = False
    nombassin = ""
    Call debut
End Sub
Private Sub debut0()
    Cb_retention.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
     bKP = False
    sval_champ = ""
    Call init_l_tab
    type_calcul = ""
    Me.Cmd_hydro.Visible = False
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    Me.Frm_bassin.Caption = "Bassin versant : "
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
    owner.fdessin.UC_graphique2.reinit 7, "Arial"
    owner.fdessin.UC_graphique2.init_title
    owner.fdessin.UC_graphique2.init_titleh ""
    owner.fdessin.UC_graphique2.init_titleb ""
    Me.Tb_bv(0).Text = "0.00"
    Me.Tb_bv(1).Text = "0"
    Me.Tb_Qf.Text = "0.00"
    Me.Tb_par(0).Text = "0.000"
    Me.Tb_par(1).Text = "0.000"
    Me.Tb_par(2).Text = "0.000"
    Me.Tb_par(3).Text = "0.000"
    Me.Tb_par(4).Text = "0"
    Call ini_ebret
'    Call reini_valeurs
    Call check_volume_enable(False)
    Me.Tb_long.Text = rempl_virgule(Format(ebret_dess.Longueur, "####0.00"))
    Me.Tb_larg.Text = rempl_virgule(Format(ebret_dess.Largeur, "####0.00"))
    Me.Tb_prof.Text = rempl_virgule(Format(ebret_dess.Profondeur, "####0.00"))
    Me.Tb_rap.Text = rempl_virgule(Format(ebret_dess.Rapport, "####0.00"))
'    Call ini_ebret
'    ebret_dess = ebret.desssret
    
    Call reini_valeurs
    Call check_volume_recup
''    If Trim(ebret_dess.type) <> "" Then
''        Call Opt_rect
''    End If
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin de rétention  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de rétention")
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
'    frm_menu.Enabled = True
    ouv_sauve = False
    Unload Frm_desprint
    Unload owner.fdessin
    owner.recharge_commentaire
End Sub


Private Sub Frm_bassin_Click()
Dim mes As String
Dim nom As String
nom = "Frm_bassin"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub

Private Sub Frm_parm_Click()
Dim mes As String
Dim nom As String
nom = "Frm_parm"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub

Private Sub Frm_rect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'owner.fdessin.UC_graphique1.Visible = False
'owner.fdessin.UC_graphique2.Visible = True

End Sub

Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub Lb_intbv_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intbv"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intpar_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intpar"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Qf_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Qf"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
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
    reponse = MsgBox("Le bassin de rétention  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de rétention")
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
    reponse = MsgBox("Le bassin de rétention  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de rétention")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_retention.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de rétention "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_retention_click
            End If
        Case Is = 7
'            Cb_retention.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de rétention "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_retention_click
            End If
    End Select
Else
'    Cb_retention.Visible = True
    frmf.Label1.Caption = "Recherche d'un bassin de rétention "
    frmf.Caption = nom
    frmf.Show 1
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_retention_click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim pict2 As New StdPicture
Dim i As Integer, nb As Integer, j As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "retention"
FrmPrint.nomobjet = Tb_titre.Text
FrmPrint.titre1 = "FICHE HYDRAULIQUE BASSIN de RETENTION"
FrmPrint.sstitre1 = "Caractéristiques " + Frm_bassin.Caption
Frm_imp.Type1 = "retention"
Frm_imp.nomobjet = Tb_titre.Text
Frm_imp.titre1 = "FICHE HYDRAULIQUE BASSIN de RETENTION"
Frm_imp.sstitre1 = "Caractéristiques " + Frm_bassin.Caption
If type_calcul = "V" Then
    FrmPrint.ssTitre2 = "Résultats intermédiaires méthode des volumes"
    Frm_imp.ssTitre2 = "Résultats intermédiaires méthode des volumes"
' certu 20080901
    FrmPrint.ssTitre2 = "Résultats intermédiaires méthode des pluies"
    Frm_imp.ssTitre2 = "Résultats intermédiaires méthode des pluies"
Else
    FrmPrint.ssTitre2 = "Résultats intermédiaires méthode hydrogramme"
    Frm_imp.ssTitre2 = "Résultats intermédiaires méthode hydrogramme"
End If
FrmPrint.ssTitre3 = ""
Frm_imp.ssTitre3 = ""
nb = (Tb_bv.count - 1) + 3  '''Tb_par.Count + 1
ReDim list_don1(nb, 3)
j = -1
For i = 0 To Tb_bv.count - 1
    j = j + 1
    list_don1(j, 1) = Lb_intbv(i).Caption
    list_don1(j, 2) = Tb_bv(i).Text
    list_don1(j, 3) = Lb_ubv(i).Caption
Next
'houpie 2004/04/06 à voir en fonction du seuil Tb_par_ep(4)
'if ebret.Seuil
'For i = 0 To Tb_par.Count - 1
Dim a As Double, b As Double, qf As Double, Ca As Double, s As Double
Dim t As Double
a = ebret.amontana
b = -ebret.bmontana
s = ebret.surface
Ca = ebret.Ca

qf = (ebret.qf / 1000) * 360 / (s * Ca / 100)
qf = qf / 60
t = (qf / (a * (b + 1))) ^ (1 / b)
If t > ebret.Seuil Then
    a = ebret.a1montana
    b = -ebret.b1montana
End If
   For i = 0 To 1
    j = j + 1
        If i = 0 Then
            list_don1(j, 1) = Lb_intpar(0).Caption + " a"
        End If
        If i = 1 Then
            list_don1(j, 1) = Lb_intpar(0).Caption + " b"
        End If
'    list_don1(j, 1) = Lb_intpar(i).Caption
    If t <= ebret.Seuil Then
        list_don1(j, 2) = Tb_par(i).Text
    Else
        list_don1(j, 2) = Tb_par(i + 2).Text
    End If
    list_don1(j, 3) = ""
Next
    j = j + 1
    list_don1(j, 1) = Lb_Qf.Caption
    list_don1(j, 2) = Tb_Qf.Text
    list_don1(j, 3) = Lb_uqf.Caption

list_int1 = rec_list(tb_resu.Text)
list_resu1 = rec_list(Tb_volume.Text)
list_resu1 = complet_list_resu1(list_resu1)
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
Set pict2 = Frm_desprint.UC_graphique2.lire_pict1()
FrmPrint.paint_picture pict1
FrmPrint.paint_picture2 pict2
SavePicture pict1, chemin_app + "dess.bmp"
SavePicture pict2, chemin_app + "dess1.bmp"
'FrmPrint.Show
Frm_imp.Show 1
End Sub
Private Function complet_list_resu1(ByVal liste1 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
i = -1
'Select Case ebstock_dess.type
'     Case Is = "circ", "cond"
'        ReDim liste(UBound(liste1) + 4, 3)
'     Case Is = "rect"
        ReDim liste(UBound(liste1) + 5, 3)
' End Select
For j = 0 To UBound(liste1)
    i = i + 1
'    ReDim Preserve liste(i, 3)
    liste(i, 1) = liste1(j, 1)
    liste(i, 2) = liste1(j, 2)
    liste(i, 3) = liste1(j, 3)
Next
'i = i + 1
''ReDim Preserve liste(i, 3)
'liste(i, 1) = ""
'liste(i, 2) = ""
'liste(i, 3) = ""
i = i + 1
'ReDim Preserve liste(i, 3)
liste(i, 1) = "Type de bassin"
liste(i, 3) = ""
'Select Case ebstock_dess.type
'     Case Is = "circ"
'     Case Is = "rect"
         liste(i, 2) = "rectangulaire"
         i = i + 1
         liste(i, 1) = Lb_intLong.Caption
         liste(i, 2) = txtVersNum(Me.Tb_long.Text)
         liste(i, 3) = Lb_uLong.Caption
         i = i + 1
         liste(i, 1) = Lb_intLarg.Caption
         liste(i, 2) = txtVersNum(Me.Tb_larg.Text)
         liste(i, 3) = Lb_uLarg.Caption
         i = i + 1
         liste(i, 1) = Lb_intprof.Caption
         liste(i, 2) = txtVersNum(Me.Tb_prof.Text)
         liste(i, 3) = Lb_uprof.Caption
         i = i + 1
         liste(i, 1) = Lb_intrap.Caption
         liste(i, 2) = txtVersNum(Me.Tb_rap.Text)
         liste(i, 3) = Lb_urap.Caption
'End Select
complet_list_resu1 = liste
End Function
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

Private Sub mnusaves_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
Call saves0
End Sub
Private Function saves0() As Boolean
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'un bassin de rétention "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du bassin de rétention " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du bassin de rétention (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
    saves0 = True
End Function

Private Sub Tb_bv_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_bv"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_bv(Index))
End Sub

Private Sub Tb_bv_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_bv"
Call sel_text(Tb_bv(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_bv_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_bv", Index, txtVersNum(Tb_bv(Index).Text))
    If Not ok Then
        Tb_bv(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_larg_Click()
Call sel_text(Tb_larg)

End Sub

Private Sub Tb_larg_GotFocus()
Call sel_text(Tb_larg)

End Sub

Private Sub Tb_larg_KeyDown(KeyCode As Integer, Shift As Integer)
    chang_larg = True
    bKP = True
    sval_champ = Tb_larg.Text
    iSels = Tb_larg.SelStart
    iSell = Tb_larg.SelLength
End Sub

Private Sub Tb_long_Click()
Call sel_text(Tb_long)

End Sub

Private Sub Tb_long_GotFocus()
Call sel_text(Tb_long)

End Sub

Private Sub Tb_long_KeyDown(KeyCode As Integer, Shift As Integer)
    chang_long = True
    bKP = True
    sval_champ = Tb_long.Text
    iSels = Tb_long.SelStart
    iSell = Tb_long.SelLength
End Sub

Private Sub Tb_par_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_par"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_par(Index))
End Sub

Private Sub Tb_par_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_par"
Call sel_text(Tb_par(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_par_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_par", Index, txtVersNum(Tb_par(Index).Text))
    If Not ok Then
        Tb_par(Index).SetFocus
        DoEvents
    Else
        Select Case Index
            Case Is = 0
                If txtVersNum(Me.Tb_par(2).Text) = 0 And ebret.a1montana = 0 Then
                    Me.Tb_par(2).Text = Me.Tb_par(0).Text
                End If
            Case Is = 1
                If txtVersNum(Me.Tb_par(3).Text) = 0 And ebret.b1montana = 0 Then
                    Me.Tb_par(3).Text = Me.Tb_par(1).Text
                End If
        End Select
    End If
    okg = True
End If

End Sub

Private Sub Tb_prof_Click()
Call sel_text(Tb_prof)

End Sub

Private Sub Tb_prof_GotFocus()
Call sel_text(Tb_prof)

End Sub

Private Sub Tb_prof_KeyDown(KeyCode As Integer, Shift As Integer)
    chang_prof = True
    bKP = True
    sval_champ = Tb_prof.Text
    iSels = Tb_prof.SelStart
    iSell = Tb_prof.SelLength
End Sub

Private Sub Tb_Qf_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Qf"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_Qf)
End Sub

Private Sub Tb_Qf_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Qf"
Call sel_text(Tb_Qf)
If change_coul Then
    Change_Couleur nom, 0
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_Qf_LostFocus()
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_Qf", -1, txtVersNum(Tb_Qf.Text))
    If Not ok Then
        Tb_Qf.SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_rap_Click()
Call sel_text(Tb_rap)

End Sub

Private Sub Tb_rap_GotFocus()
Call sel_text(Tb_rap)

End Sub

Private Sub Tb_rap_KeyDown(KeyCode As Integer, Shift As Integer)
    chang_rap = True
     bKP = True
     sval_champ = Tb_rap.Text
    iSels = Tb_rap.SelStart
    iSell = Tb_rap.SelLength
End Sub

Private Sub Tb_resu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'owner.fdessin.UC_graphique1.Visible = True
'owner.fdessin.UC_graphique2.Visible = False

End Sub

Private Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    Call save0
End Sub
Public Function save0()
Dim ok As Boolean
ok = False

    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        ok = save(False)
    Else
        ok = saves0
    End If
save0 = ok
End Function
Public Function save(ByVal bsous As Boolean) As Boolean
Dim za As st_savret
Dim za1 As st_savret1
Dim i As Integer, isave As Integer
Dim reponse As Integer
 
'Dim ret_sauve As Boolean
'ret_sauve = False
If Trim(Tb_titre.Text) <> "" Then
    Call funlockb
'   ebret.nombv = ebv.nom 'nombassin
   ebret.nombv = nombassin
    ebret.nom = nombassin 'ebv.nom
    ebret.type_calcul = type_calcul
    ebret.desssret = ebret_dess
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            za = za1.stsavret
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un bassin de rétention")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "retention"
            za.nom = Trim(Tb_titre.Text)
            za.retention = ebret
            za1.stsavret = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        za.type = "retention"
        za.nom = Tb_titre.Text
        za.retention = ebret
        za1.stsavret = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        ret_texte = Trim(Tb_titre.Text)
        Cb_retention.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom du bassin de rétention n'est pas renseigné.", , "Sauvegarde d'un bassin de rétention")
End If
'save = ret_sauve
    save = True
 
Exit Function
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If

Call flockb(nom_fich)
End Function
Private Sub mnusuppr_Click()
Dim za As st_savret
Dim za1 As st_savret1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_retention.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_retention.Text) + " va être supprimé .", 4, "Suppression d'un bassin de rétention")
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
            za = za1.stsavret
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_retention.Text)) Then
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
    Close #lhFicDbf1
    Kill nom
    Call flockb(nom_fich)
    Call lect_fich
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
    Call reini_valeurs
    Call ini_ebret
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique2.graphique_clear
    owner.fdessin.UC_graphique1.init_titleb ""
    owner.fdessin.UC_graphique2.init_titleb ""
    Me.Frm_bassin.Caption = "Bassin versant : "
'Houpie 2005/03/21
'    Me.Tb_bv(0).Text = "0"
    Me.Tb_bv(0).Text = "0.00"
    Me.Tb_bv(1).Text = "0"
    Me.Tb_Qf.Text = "0.00"
    Me.Tb_par(0).Text = "0.000"
    Me.Tb_par(1).Text = "0.000"
    Me.Tb_par(2).Text = "0.000"
    Me.Tb_par(3).Text = "0.000"
    Me.Tb_par(4).Text = "0"
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

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Private Sub Cmd_calcul_Click()
    Dim mes As String
    Dim nom As String
    Dim reponse As Integer
    nom = "Cmd_calcul"
    mes = Rec_Mes(nom, 0)
    Change_Focus nom, 0
    owner.affich_aide Me.Name, mes
    reponse = MsgBox("Attention!Pour valider ce calcul,vérifier que la durée Tmax appartient bien au domaine de validité de vos coefficients a et b.", 1, "Calcul par la méthode des pluies")
If reponse = 1 Then
'        2=annuler,ok=1
    Call key13(Me)
    type_calcul = "V"
    Call ini_tbresu
    Call Calc_volume
    Call reini_dess
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
End If
End Sub
Private Sub Calc_volume()
Dim delta1 As Double, deltaH As Double, v As Double
Dim qvm() As Variant
Dim ebcourbe As courbe_dess
'calcul de deltah
Dim a As Double, b As Double, qf As Double, Ca As Double, s As Double
Dim t As Double, h As Double, hf As Double, dtot As Double
Dim sresult As String, sresult1 As String
a = ebret.amontana
b = -ebret.bmontana
s = ebret.surface
Ca = ebret.Ca

qf = (ebret.qf / 1000) * 360 / (s * Ca / 100)
qf = qf / 60
If b <= -1 Then
        MsgBox "Erreur dans les coefficients de Montana ..", vbExclamation, "Méthode des volumes"
Else
t = (qf / (a * (b + 1))) ^ (1 / b)
If t > ebret.Seuil Then
    a = ebret.a1montana
    b = -ebret.b1montana
    t = (qf / (a * (b + 1))) ^ (1 / b)

End If
h = a * t ^ (b + 1)
hf = qf * t
Call modi_tbresu
sresult = " Calcul du volume maximum stocké "
sresult = sresult + Chr(13) + Chr(10) + "  Durée   = " + ajout_zero(Trim(str(Round(t, 2)))) + " mn"
sresult = sresult + Chr(13) + Chr(10) + "  Hauteur de pluie   = " + ajout_zero(Trim(str(Round(h, 2)))) + " mm"
sresult = sresult + Chr(13) + Chr(10) + "  Hauteur de fuite   = " + ajout_zero(Trim(str(Round(hf, 2)))) + " mm"
ebret_dess.duree = Round(t, 2)
ebret_dess.Hpluie = Round(h, 2)
ebret_dess.Hfuite = Round(hf, 2)
If qf > 0 Then
    delta1 = (qf / (a * (b + 1))) ^ (1 / b)
    deltaH = delta1 * ((-b * qf) / (b + 1))
'calcul du volume

    v = 10 * s * (Ca / 100) * deltaH
    ebret_dess.Hpluie = Round(10 * s * (Ca / 100) * h, 2)
    ebret_dess.Hfuite = Round(10 * s * (Ca / 100) * hf, 2)

    sresult = sresult + Chr(13) + Chr(10) + "  deltah   = " + ajout_zero(Trim(str(Round(deltaH, 3)))) + " mm"
    ebret.deltaH = Round(deltaH, 3)
    
    sresult = sresult + Chr(13) + Chr(10) + Chr(10) + " Volume ruisselé   = " + ajout_zero(Trim(str(Round(ebret_dess.Hpluie, 3)))) + " m3"
    sresult = sresult + Chr(13) + Chr(10) + "  Volume évacué   = " + ajout_zero(Trim(str(Round(ebret_dess.Hfuite, 3)))) + " m3"
    sresult1 = "  Volume de stockage   = " + ajout_zero(Trim(str(Round(v, 3)))) + " m3"
    ebret.volume = Round(v, 3)
'    Me.Lb_resu.Caption = sresult
    ebcourbe.duree = ebret_dess.duree
    ebcourbe.quantite = ebret.qf
    ebcourbe.volume = ebret.volume
    ebcourbe.hauteur = ebret_dess.Hpluie
Call init_graph_courbe(owner.fdessin.UC_graphique1, ebcourbe)
dtot = ((Int((ebret_dess.duree * 1.5) / 10) + 1) * 10) / 100#
ReDim qvm(101, 2)
For i = 1 To UBound(qvm)
qvm(i, 1) = (i - 1) * dtot
qvm(i, 2) = 10 * s * (Ca / 100) * a * (dtot * (i - 1)) ^ (b + 1)
Next
owner.fdessin.UC_graphique1.init_titleb "Construction Graphique"
owner.fdessin.UC_graphique1.init_lbvh "m3"
owner.fdessin.UC_graphique1.init_lbhd "mn"
Frm_imp.des2_titrb = "Construction Graphique"
FrmPrint.des2_titrb = "Construction Graphique"
owner.fdessin.UC_graphique1.dess_poly qvm, "N", couleur.bleu, 1
owner.fdessin.UC_graphique1.dess_lign_point ebret_dess.duree, 0, ebret_dess.duree, ebret_dess.Hfuite, couleur.magenta
owner.fdessin.UC_graphique1.dess_lign ebret_dess.duree, ebret_dess.Hfuite, ebret_dess.duree, ebret_dess.Hpluie, couleur.magenta, 2
'*********dessin fenêtre tampon
Call init_graph_courbe(Frm_desprint.UC_graphique2, ebcourbe)
Frm_desprint.UC_graphique2.init_lbvh "m3"
Frm_desprint.UC_graphique2.init_lbhd "mn"
Frm_desprint.UC_graphique2.dess_poly qvm, "N", couleur.bleu, 1
Frm_desprint.UC_graphique2.dess_lign_point ebret_dess.duree, 0, ebret_dess.duree, ebret_dess.Hfuite, couleur.magenta
Frm_desprint.UC_graphique2.dess_lign ebret_dess.duree, ebret_dess.Hfuite, ebret_dess.duree, ebret_dess.Hpluie, couleur.magenta, 2
'***************
End If
    ouv_sauve = True
'    Me.Cmd_calcul.Enabled = False
   Me.tb_resu.Text = sresult
   Me.Tb_volume.Text = sresult1
    Me.Cmd_graph.Enabled = True
End If ' fin erreur
End Sub
Private Sub meAffiche()
    DoEvents
    Me.Show
End Sub
Public Sub calc_hydro()
Dim pas As Integer, tt As Integer
Dim vt As Double
Dim sresult As String, sresult1 As String
''ehyd <-
On Error GoTo test_Error
pas = ehyd.pas
        tt = calcul_hyeto(ehyd, pas)
 '       Call dessin_hyeto1
'        ' le hyeto brut est stocké dabs la table globale hpluie()
'
'' ebv ,ehyd <-variable globale
        Call calcul_hydro(pas)  ' en sortie Q
          Call dessine_hydro   '(utilise Q, Hpluie,ehyd)
'Call modi_tbresu
''modifs WAGNER 24/05/2004
'vt = calc_stock(ehyd)
'    sresult = " Calcul des valeurs au volume maximum stocké "
'    sresult = sresult + Chr(13) + Chr(10) + " Volume ruisselé   = " + ajout_zero(Trim(Str(vt))) + " m3"
'    sresult = sresult + Chr(13) + Chr(10) + "  Volume évacué   = " + ajout_zero(Trim(Str(vt - ehyd.vstock))) + " m3"
'    sresult1 = "  Volume de stockage   = " + ajout_zero(Trim(Str(ehyd.vstock))) + " m3"

''modifs WAGNER 24/05/2004
sresult = calc_stock(ehyd)
vt = val(stGetToken(sresult, ","))
Tdeb = RTrim(stGetToken(sresult, ","))
Tfin = RTrim(stGetToken(sresult, ","))
    sresult = " Fonctionnement du bassin "
   sresult = sresult + Chr(13) + Chr(10) + " Début du stockage   = " + ajout_zero(Trim(str(Tdeb))) + " mn"
    sresult = sresult + Chr(13) + Chr(10) + " Fin du stockage   = " + ajout_zero(Trim(str(Tfin))) + " mn"
    sresult1 = "  Volume de stockage   = " + ajout_zero(Trim(str(ehyd.vstock))) + " m3"
''fin modifs
    
    ebret.volume = ehyd.vstock
    Me.tb_resu.Text = sresult
   Me.Tb_volume.Text = sresult1
     Me.Cmd_graph.Enabled = True
   ouv_sauve = True
  Exit Sub
test_Error:
        Call print_erreur("anomalie dans calc_hydro ")
      

End Sub
Private Function calc_stock(ByRef ehyd As st_hydr) As String
'Private Function calc_stock(ByRef ehyd As st_hydr) As Double
Dim i As Integer
Dim qpas As Double, pas As Integer
Dim vstock As Double, vtomb As Double, qfuite As Double
Dim max_atteint As Boolean
'modifs WAGNER 24/05/2004
Dim vdestock As Double, vstockmax As Double
Dim Tdeb As Integer, Tfin As Integer
'fin modifs
On Error GoTo test_Error
max_atteint = False
pas = ehyd.pas
vstock = 0
vtomb = 0
qfuite = ehyd.qfuite
''modifs WAGNER 24/05/2004
'For i = 1 To UBound(Q)
'qpas = Q(i)
'    If qpas > qfuite / 1000 Then
'        max_atteint = True
'        vstock = vstock + (qpas - qfuite / 1000) * pas * 60
'        ehyd.vstock = Round(vstock, 3)
'        If vstock > 0 Then
'            vtomb = vtomb + qpas * pas * 60
'        End If
'    Else
'        If Not max_atteint Then
'            vtomb = vtomb + qpas * pas * 60
'        End If
'    End If
'Next
'ehyd.vstock = Round(vstock, 3)
' calc_stock = Round(vtomb, 3)
''modifs WAGNER 24/05/2004
vstockmax = 0
Tdeb = 0
Tfin = 0
 For i = 1 To UBound(Q)
qpas = Q(i)
vtomb = vtomb + qpas * pas * 60
    If qpas > qfuite / 1000 Then
        If Tdeb = 0 Then
            Tdeb = i * pas
        End If
        vstock = vstock + (qpas - qfuite / 1000) * pas * 60
        ehyd.vstock = Round(vstock, 3)
    Else
        If vstock > 0 Then
            vdestock = (qfuite / 1000 - qpas) * pas * 60
            If vdestock < vstock Then
                vstock = vstock - vdestock
                Else
               Tfin = i * pas
                vstock = 0
            End If
        End If
    End If
    If vstock > vstockmax Then
        vstockmax = vstock
    End If
Next
qpas = 0
        While vstock > 0
            vdestock = (qfuite / 1000 - qpas) * pas * 60
            If vdestock < vstock Then
                vstock = vstock - vdestock
                i = i + 1
                Else
               Tfin = i * pas
                vstock = 0
            End If
        Wend
Tdeb = Tdeb - pas
c$ = str(vtomb) + "," + str(Tdeb) + "," + str(Tfin)
 ehyd.vstock = Round(vstockmax, 3)
 calc_stock = c$
''fin modifs
   Exit Function
test_Error:
        Call print_erreur("anomalie dans calc_stock")

End Function
Function stGetToken(stLn$, stDelim$) As String
    On Error GoTo GetTokenError

    iOpenQuote% = InStr(1, stLn$, """")
    iDelim% = InStr(1, stLn$, stDelim$)

    If (iOpenQuote% > 0) And (iOpenQuote% < iDelim%) Then
         iCloseQuote% = InStr(iOpenQuote% + 1, stLn$, """")
         iDelim% = InStr(iCloseQuote% + 1, stLn$, stDelim$)
    End If

    If (iDelim% <> 0) Then
         stToken$ = LTrim$(RTrim$(Mid$(stLn$, 1, iDelim% - 1)))
         stLn$ = Mid$(stLn$, iDelim% + 1)
    Else
         stToken$ = LTrim$(RTrim$(Mid$(stLn$, 1)))
         stLn$ = ""
    End If

    If (Len(stToken$) > 0) Then
         If (Mid$(stToken$, 1, 1) = """") Then
              stToken$ = Mid$(stToken$, 2)
         End If
         If (Mid$(stToken$, Len(stToken$), 1) = """") Then
              stToken$ = Mid$(stToken$, 1, Len(stToken$) - 1)
         End If
    End If
    stGetToken = stToken$

GetTokenExit:
    Exit Function

GetTokenError:
    Resume GetTokenExit
End Function

Private Sub dessine_hydro()
Dim q10() As Variant, q11() As Variant
Dim i As Integer
Dim vt As Double
Dim ok As Boolean
ReDim q0(UBound(Q))
ReDim q10(UBound(Q), 2)
ReDim q11(UBound(Q), 2)
On Error GoTo test_Error
ok = True
vt = 0
For i = 1 To UBound(Q)
    q10(i, 1) = i * ehyd.pas * 1#
    q10(i, 2) = Q(i) * 1000#
' dessin de la fuite
    q11(i, 1) = i * ehyd.pas
    If ehyd.qfuite < q10(i, 2) Then
        ok = False
    End If
    If ehyd.qfuite > q10(i, 2) + vt Then
        q11(i, 2) = q10(i, 2) + vt
        vt = 0
        
    Else
        vt = vt + q10(i, 2) - ehyd.qfuite
        q11(i, 2) = ehyd.qfuite
    End If
Next
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb "HYDROGRAMME DE RUISSELLEMENT"
    owner.fdessin.UC_graphique1.init_arrondi_y 1
    owner.fdessin.UC_graphique1.init_MaxYn q10
    owner.fdessin.UC_graphique1.init_EchYn 1# '0.6
    owner.fdessin.UC_graphique1.init_MaxXn q10
    owner.fdessin.UC_graphique1.init_EchXn 1#
    owner.fdessin.UC_graphique1.dess_cadre 8, 2, 50, 0, 0, 0, 6, 1, 10
    owner.fdessin.UC_graphique1.init_lbvh "l/s"
    owner.fdessin.UC_graphique1.init_lbhd "mn"
    owner.fdessin.UC_graphique1.dess_courbe q10, "N", &H80FF80
' dessin de la fuite
    owner.fdessin.UC_graphique1.dess_courbe q11, "N", &H80C0FF
'*******dessin fenêtre tampon
    Frm_desprint.UC_graphique2.reinit 7, "Arial"
    Frm_desprint.UC_graphique2.graphique_clear
    Frm_desprint.UC_graphique2.init_title
    Frm_desprint.UC_graphique2.init_titleh ""
    Frm_desprint.UC_graphique2.init_titleb "HYDROGRAMME DE RUISSELLEMENT"
    Frm_imp.des2_titrb = "HYDROGRAMME DE RUISSELLEMENT"
    FrmPrint.des2_titrb = "HYDROGRAMME DE RUISSELLEMENT"
    Frm_desprint.UC_graphique2.init_arrondi_y 1
    Frm_desprint.UC_graphique2.init_MaxYn q10
    Frm_desprint.UC_graphique2.init_EchYn 1# '0.6
    Frm_desprint.UC_graphique2.init_MaxXn q10
    Frm_desprint.UC_graphique2.init_EchXn 1#
    Frm_desprint.UC_graphique2.dess_cadre 8, 2, 50, 0, 0, 0, 6, 1, 10
    Frm_desprint.UC_graphique2.init_lbvh "l/s"
    Frm_desprint.UC_graphique2.init_lbhd "mn"
    Frm_desprint.UC_graphique2.dess_courbe q10, "N", &H80FF80
' dessin de la fuite
    Frm_desprint.UC_graphique2.dess_courbe q11, "N", &H80C0FF
'****************************
   Exit Sub
test_Error:
        Call print_erreur("anomalie dans dessine_hydro")

End Sub
Public Sub ini_debit(ByVal nom As String)
    nombassin = nom
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphique1.graphique_clear
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
    owner.fdessin.UC_graphique2.reinit 7, "Arial"
    owner.fdessin.UC_graphique2.init_title
    owner.fdessin.UC_graphique2.init_titleh ""
    owner.fdessin.UC_graphique2.init_titleb ""
    ehyd.qfuite = 0
    ebv.qfuite = 0
'julienne 2001/12/12
'    Me.Tb_titre.Text = ""
'    Me.Cb_retention.Text = ""
   If Trim(ebv.Qchoisi) <> "" Then
        Me.Frm_bassin.Caption = "Bassin versant : " + Trim(nombassin) 'Trim(ebv.nom)
'Houpie 2005/03/21
        Me.Tb_bv(0).Text = rempl_virgule(Format(ebv.surface, "###0.00"))
'        Me.Tb_bv(0).Text = ajout_zero(Trim(Str(ebv.surface)))
        Me.Tb_bv(1).Text = rempl_virgule(Format(ebv.imper, "###0"))
'        Me.Tb_Qf.Text = rempl_virgule(Format(ebv.qfuite, "###0"))
        Me.Tb_par(0).Text = rempl_virgule(Format(eph.amontana, "##0.000"))
        Me.Tb_par(1).Text = rempl_virgule(Format(eph.bmontana, "##0.000"))
        Me.Tb_par(2).Text = rempl_virgule(Format(eph.a1montana, "##0.000"))
        Me.Tb_par(3).Text = rempl_virgule(Format(eph.b1montana, "##0.000"))
        Me.Tb_par(4).Text = rempl_virgule(Format(eph.Seuil, "####0"))
        Me.Cmd_hydro.Visible = True
''       Call reini_valeurs
        nombassin = nom
        Me.Frm_bassin.Caption = "Bassin versant : " + Trim(nombassin)  ' Trim(ebv.nom)
    Else
        Me.Frm_bassin.Caption = "Bassin versant : "
'Houpie 2005/03/21
'        Me.Tb_bv(0).Text = "0"
        Me.Tb_bv(0).Text = "0.00"
        Me.Tb_bv(1).Text = "0"
        Me.Tb_Qf.Text = "0.00"
        Me.Tb_par(0).Text = "0.000"
        Me.Tb_par(1).Text = "0.000"
        Me.Tb_par(2).Text = "0.000"
        Me.Tb_par(3).Text = "0.000"
        Me.Tb_par(4).Text = "0"
''        Call reini_valeurs

   End If
' Call ini_ebret_dess
    Call init_form_dess
   Call reini_valeurs
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Public Sub reini_valeurs()
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique2.graphique_clear
    owner.fdessin.UC_graphique1.init_titleb ""
    owner.fdessin.UC_graphique2.init_titleb ""
    Call ini_tbresu
    If ebret.amontana > 0 And ebret.bmontana > 0 And ebret.a1montana > 0 _
        And ebret.b1montana > 0 And ebret.Seuil > 0 And ebret.surface > 0 _
        And ebret.Ca > 0 And ebret.qf > 0 Then
        Me.Cmd_calcul.Enabled = True
        Me.Cmd_hydro.Enabled = True
        ' impression vraie
                    Me.mnuprint.Enabled = True
    Else
        Me.Cmd_calcul.Enabled = False
        Me.Cmd_hydro.Enabled = False
        ' impression false
                    Me.mnuprint.Enabled = False
    End If
        Me.Frm_rect.Visible = False
        Me.Cmd_graph.Enabled = False
        Me.Cmd_schema.Enabled = False
        ouv_sauve = True

End Sub

Private Sub ini_tbresu()
'    Me.tb_resu.BackColor = &H8000000B
'    Me.tb_resu.BorderStyle = 1
    Me.tb_resu.Text = ""
'    Me.Tb_volume.BackColor = &H8000000B
'    Me.Tb_volume.BorderStyle = 1
    Me.Tb_volume.Text = ""
End Sub
Private Sub modi_tbresu()
'    Me.tb_resu.BackColor = &H80000009
'    Me.tb_resu.BorderStyle = 1
'    Me.Tb_volume.BackColor = &H80000009
'    Me.Tb_volume.BorderStyle = 1
End Sub
Public Sub ini_ebret()
    ebret.nom = ""
    ebret.surface = 0
    ebret.Ca = 0
    ebret.qf = 0
    ebret.amontana = 0#
    ebret.bmontana = 0#
    ebret.deltaH = 0#
    ebret.volume = 0#
    ebret.a1montana = 0#
    ebret.b1montana = 0#
    ebret.Seuil = 0
    Call ini_ebret_dess
    ebret.desssret = ebret_dess
End Sub
Public Sub ini_ebret_dess()
    ebret_dess.type = " "
    ebret_dess.coef = 0
    ebret_dess.Longueur = 0#
    ebret_dess.Largeur = 0#
    ebret_dess.Profondeur = 0#
    ebret_dess.Rapport = 0#
    ebret_dess.opt_long = False
    ebret_dess.opt_larg = False
    ebret_dess.opt_prof = False
    ebret_dess.opt_rap = False
End Sub
Private Sub reini_dess()

' Call check_volume_enable(False)
   Me.Frm_rect.Visible = True
'    Me.Tb_long.Text = ""
'    Me.Tb_larg.Text = ""
'    Me.Tb_prof.Text = ""
'    Me.Tb_rap.Text = ""
'    ebret_dess = ebret.desssret
    Call check_volume_recup
    Call init_form_dess
'    chang_diam = True
'    Me.Tb_diam.Text = Format(ebstock.dessstock.Diametre, "####0.00")
'    Me.Tb_long.Text = Format(ebstock.dessstock.Longueur, "####0.00")
'    Me.Tb_larg.Text = Format(ebstock.dessstock.Largeur, "####0.00")
'    Me.Tb_prof.Text = Format(ebstock.dessstock.Profondeur, "####0.00")
'    Me.Tb_rap.Text = Format(ebstock.dessstock.Rapport, "####0.00")

End Sub
Private Sub check_volume_enable(ByVal ok As Boolean)
 Me.Chk_larg.Enabled = ok
  Me.Chk_long.Enabled = ok
  Me.Chk_prof.Enabled = ok
  Me.Chk_rap.Enabled = ok
  Call check_volume_saisie
End Sub
Private Sub check_volume_saisie()

    Me.Tb_larg.Enabled = Me.Chk_larg.Enabled
    Me.Tb_long.Enabled = Me.Chk_long.Enabled
    Me.Tb_prof.Enabled = Me.Chk_prof.Enabled
    Me.Tb_rap.Enabled = Me.Chk_rap.Enabled
'    ebret_dess.opt_larg = (Me.Chk_larg.Value = 1)
'    ebret_dess.opt_long = (Me.Chk_long.Value = 1)
'    ebret_dess.opt_prof = (Me.Chk_prof.Value = 1)
'    ebret_dess.opt_rap = (Me.Chk_rap.Value = 1)
End Sub
Private Sub Chk_long_Click()
If Me.Chk_long.Enabled Then
    Call check_volume
End If
End Sub

Private Sub Chk_prof_Click()
If Me.Chk_prof.Enabled Then
    Call check_volume
End If
End Sub

Private Sub Chk_rap_Click()
If Me.Chk_rap.Enabled Then
    Call check_volume
End If
End Sub
Private Sub init_form_dess()
Call check_volume_enable(False)
Chk_larg.Value = 0
Chk_long.Value = 0
Chk_prof.Value = 0
Chk_rap.Value = 0
 '   ebret_dess = ebret.desssret

    Me.Tb_long.Text = rempl_virgule(Format(ebret_dess.Longueur, "####0.00"))
    Me.Tb_larg.Text = rempl_virgule(Format(ebret_dess.Largeur, "####0.00"))
    Me.Tb_prof.Text = rempl_virgule(Format(ebret_dess.Profondeur, "####0.00"))
    Me.Tb_rap.Text = rempl_virgule(Format(ebret_dess.Rapport, "####0.00"))
    Call check_volume_recup
    Call check_volume
'    Call check_volume_enable(True)
    Call Opt_rect
'    If Trim(ebret_dess.type) <> "" Then
 '       Me.Frm_rect.Visible = True
'        If ebstock_dess.type = "circ" Then
'            Opt_cir.Value = True
'            Opt_rect.Value = False
'            Call Opt_cir_Click
'
'        Else
'            Opt_cir.Value = False
'            Opt_rect.Value = True
'            Call Opt_rect_Click
'        End If
'        Call dess_stock
'    End If

End Sub
Private Sub check_volume()
If Me.Chk_larg.Value + Me.Chk_long.Value + Me.Chk_prof.Value + Me.Chk_rap.Value = 2 Then

If Me.Chk_larg.Value = 0 Then
    Me.Chk_larg.Enabled = False
End If
If Me.Chk_long.Value = 0 Then
    Me.Chk_long.Enabled = False
End If
If Me.Chk_prof.Value = 0 Then
    Me.Chk_prof.Enabled = False
End If
If Me.Chk_rap.Value = 0 Then
    Me.Chk_rap.Enabled = False
End If
Call check_volume_saisie
Call calcul_dimension
Else
 Call check_volume_enable(True)
ebret_dess.opt_larg = (Me.Chk_larg.Value = 1)
ebret_dess.opt_long = (Me.Chk_long.Value = 1)
ebret_dess.opt_prof = (Me.Chk_prof.Value = 1)
ebret_dess.opt_rap = (Me.Chk_rap.Value = 1)
'  Call check_volume_saisie
End If
End Sub
Private Sub calcul_dimension()
Dim surf As Double, xlon As Double, xlar As Double, haut As Double, rap As Double
owner.fdessin.UC_graphique2.graphique_clear
' impression false
Me.mnuprint.Enabled = False
ebret_dess.opt_larg = (Me.Chk_larg.Value = 1)
    ebret_dess.opt_long = (Me.Chk_long.Value = 1)
    ebret_dess.opt_prof = (Me.Chk_prof.Value = 1)
    ebret_dess.opt_rap = (Me.Chk_rap.Value = 1)
xlon = ebret_dess.Longueur
xlar = ebret_dess.Largeur
haut = ebret_dess.Profondeur
rap = ebret_dess.Rapport
If xlon + xlar + haut + rap > 0 Then
If ebret_dess.opt_long And ebret_dess.opt_larg Then
    If xlon > 0 And xlar > 0 Then
        surf = ebret.volume / xlon
        haut = surf / xlar
        rap = xlar / haut
    Else
        haut = 0#
        rap = 0#
    End If
End If
If ebret_dess.opt_long And ebret_dess.opt_prof Then
    If xlon > 0 And haut > 0 Then
        surf = ebret.volume / xlon
        xlar = surf / haut
        rap = xlar / haut
    Else
        xlar = 0#
        rap = 0#
    End If
End If
If ebret_dess.opt_long And ebret_dess.opt_rap Then
    If xlon > 0 And rap > 0 Then
        surf = ebret.volume / xlon
        haut = Sqr(surf / rap)
        xlar = rap * haut
    Else
        haut = 0#
        xlar = 0#
    End If
End If
If ebret_dess.opt_larg And ebret_dess.opt_prof Then
    If xlar > 0 And haut > 0 Then
        surf = xlar * haut
        xlon = ebret.volume / surf
        rap = xlar / haut
    Else
        xlon = 0#
        rap = 0#
    End If
End If
If ebret_dess.opt_larg And ebret_dess.opt_rap Then
    If xlar > 0 And rap > 0 Then
        haut = xlar / rap
        surf = xlar * haut
        xlon = ebret.volume / surf
    Else
        haut = 0#
        xlon = 0#
    End If
End If
If ebret_dess.opt_prof And ebret_dess.opt_rap Then
    If haut > 0 And rap > 0 Then
        xlar = haut * rap
        surf = xlar * haut
        xlon = ebret.volume / surf
    Else
        xlar = 0#
        xlon = 0#
    End If
End If
If Not bklong Then
Me.Tb_long.Text = rempl_virgule(Format(Round(xlon, 2), "##0.00"))
End If
If Not bklarg Then
Me.Tb_larg.Text = rempl_virgule(Format(Round(xlar, 2), "##0.00"))
End If
If Not bkprof Then
Me.Tb_prof.Text = rempl_virgule(Format(Round(haut, 2), "##0.00"))
End If
If Not bkrap Then
Me.Tb_rap.Text = rempl_virgule(Format(Round(rap, 2), "##0.00"))
End If
'bklong = False
'bklarg = False
'bkprof = False
'bkrap = False

ebret_dess.Longueur = Round(xlon, 2)
ebret_dess.Largeur = Round(xlar, 2)
ebret_dess.Profondeur = Round(haut, 2)
ebret_dess.Rapport = Round(rap, 2)
Call dess_stock

End If
End Sub
Private Sub check_volume_recup()
        If ebret_dess.opt_long Then
        Me.Chk_long.Value = 1
        Else
        Me.Chk_long.Value = 0
    End If
    If ebret_dess.opt_larg Then
        Me.Chk_larg.Value = 1
         Else
        Me.Chk_larg.Value = 0
   End If
    If ebret_dess.opt_prof Then
        Me.Chk_prof.Value = 1
        Else
        Me.Chk_prof.Value = 0
    End If
    If ebret_dess.opt_rap Then
        Me.Chk_rap.Value = 1
    Else
        Me.Chk_rap.Value = 0
    End If
    Me.Chk_larg.Enabled = ebret_dess.opt_larg
    Me.Chk_long.Enabled = ebret_dess.opt_long
    Me.Chk_prof.Enabled = ebret_dess.opt_prof
    Me.Chk_rap.Enabled = ebret_dess.opt_rap

    Me.Tb_larg.Enabled = Me.Chk_larg.Enabled
    Me.Tb_long.Enabled = Me.Chk_long.Enabled
    Me.Tb_prof.Enabled = Me.Chk_prof.Enabled
    Me.Tb_rap.Enabled = Me.Chk_rap.Enabled
End Sub

Private Sub Tb_bv_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
'Houpie 2005/03/21
'                nom = verif_cart0(Tb_bv(Index).Text, "Saisie de la surface du B.V.", "I")
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie de la surface du B.V.", "R")
            Case Is = 1
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie du coefficient d'apport du B.V.", "I")
        End Select
  If nom = "" Then
    Tb_bv(Index).Text = sval_champ
    Tb_bv(Index).SelStart = iSels
    Tb_bv(Index).SelLength = iSell
  End If
End If
'****

Select Case Index
    Case Is = 0
        ebret.surface = txtVersNum(Me.Tb_bv(0).Text)
    Case Is = 1
        ebret.Ca = txtVersNum(Me.Tb_bv(1).Text)
End Select
'Call ini_ebret_dess
    Me.Frm_bassin.Caption = "Bassin versant : "
    nombassin = ""
    Me.Cmd_hydro.Visible = False
    Call reini_valeurs
     sval_champ = ""
    bKP = False
End Sub

Private Sub Tb_bv_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_bv(Index).Text
    iSels = Tb_bv(Index).SelStart
    iSell = Tb_bv(Index).SelLength
'    If Len(Tb_bv(Index).Text) <= Tb_bv(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie de la surface du B.V.", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie du coefficient d'apport du B.V.", "I")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_larg_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String
bklarg = False
If bKP Then
        nom = verif_cart0(Tb_larg.Text, "Saisie de la largeur", "R")
  If nom = "" Then
    Tb_larg.Text = sval_champ
    Tb_larg.SelStart = iSels
    Tb_larg.SelLength = iSell
  Else
'  End If
'End If
'****
bklarg = True

xlar = txtVersNum(Tb_larg.Text)
ebret_dess.Largeur = Round(xlar, 2)
If Chk_larg.Value = 0 And Chk_larg.Enabled Then
    Chk_larg.Value = 1
End If
Cmd_schema.Enabled = False
bklarg = True

If chang_larg Then
    chang_larg = False
    If Me.Chk_larg.Value + Me.Chk_long.Value + Me.Chk_prof.Value + Me.Chk_rap.Value = 2 Then
        Call calcul_dimension
    End If
'    Call dess_stock
End If
  End If
End If

 sval_champ = ""
bklarg = False

 bKP = False

End Sub
Private Sub Tb_larg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    chang_larg = True
    bKP = True
    sval_champ = Tb_larg.Text
    iSels = Tb_larg.SelStart
    iSell = Tb_larg.SelLength
'    If Len(Tb_larg.Text) <= Tb_larg.MaxLength Then
'        KeyAscii = verif_car(Tb_larg.Text, KeyAscii, "Saisie de la largeur", "R")
'    End If
End If
End Sub

Private Sub Tb_long_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String
bklong = False

If bKP Then
        nom = verif_cart0(Tb_long.Text, "Saisie de la longueur", "R")
  If nom = "" Then
    Tb_long.Text = sval_champ
    Tb_long.SelStart = iSels
    Tb_long.SelLength = iSell
  Else
'  End If
'End If
'****

bklong = True

xlon = txtVersNum(Tb_long.Text)
ebret_dess.Longueur = Round(xlon, 2)
If Chk_long.Value = 0 And Chk_long.Enabled Then
    Chk_long.Value = 1
End If
bklong = True
Cmd_schema.Enabled = False
If chang_long Then
    chang_long = False
    If Me.Chk_larg.Value + Me.Chk_long.Value + Me.Chk_prof.Value + Me.Chk_rap.Value = 2 Then
        Call calcul_dimension
    End If
'    Call dess_stock
End If
  End If
End If

 sval_champ = ""
 bKP = False
bklong = False

End Sub
Private Sub Tb_long_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    chang_long = True
    bKP = True
    sval_champ = Tb_long.Text
    iSels = Tb_long.SelStart
    iSell = Tb_long.SelLength
'    If Len(Tb_long.Text) <= Tb_long.MaxLength Then
'        KeyAscii = verif_car(Tb_long.Text, KeyAscii, "Saisie de la longueur", "R")
'    End If
End If
End Sub

Private Sub tb_par_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_par(Index).Text, "Saisie du coefficient a1 de Montana", "R")
            Case Is = 1
                nom = verif_cart0(Tb_par(Index).Text, "Saisie du coefficient b1 de Montana", "R")
            Case Is = 2
                nom = verif_cart0(Tb_par(Index).Text, "Saisie du coefficient a2 de Montana", "R")
            Case Is = 3
                nom = verif_cart0(Tb_par(Index).Text, "Saisie du coefficient b2 de Montana", "R")
            Case Is = 4
                nom = verif_cart0(Tb_par(Index).Text, "Saisie du seuil", "I")
        End Select
  If nom = "" Then
    Tb_par(Index).Text = sval_champ
    Tb_par(Index).SelStart = iSels
    Tb_par(Index).SelLength = iSell
  End If
End If
'****

Select Case Index
    Case Is = 0
        ebret.amontana = txtVersNum(Me.Tb_par(0).Text)
    Case Is = 1
        ebret.bmontana = txtVersNum(Me.Tb_par(1).Text)
    Case Is = 2
        ebret.a1montana = txtVersNum(Me.Tb_par(2).Text)
    Case Is = 3
        ebret.b1montana = txtVersNum(Me.Tb_par(3).Text)
    Case Is = 4
        ebret.Seuil = txtVersNum(Me.Tb_par(4).Text)
End Select
    Me.Frm_bassin.Caption = "Bassin versant : "
    nombassin = ""
    Me.Cmd_hydro.Visible = False
    Call reini_valeurs
     sval_champ = ""
    bKP = False
End Sub

Private Sub Tb_par_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_par(Index).Text
    iSels = Tb_par(Index).SelStart
    iSell = Tb_par(Index).SelLength
'    If Len(Tb_par(Index).Text) <= Tb_par(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_par(Index).Text, KeyAscii, "Saisie du paramètre a de Montana", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_par(Index).Text, KeyAscii, "Saisie du paramètre b de Montana", "R")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_prof_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String
bkprof = False
If bKP Then
        nom = verif_cart0(Tb_prof.Text, "Saisie de la hauteur d'eau", "R")
  If nom = "" Then
    Tb_prof.Text = sval_champ
    Tb_prof.SelStart = iSels
    Tb_prof.SelLength = iSell
  Else
'  End If
'End If
'****
bkprof = True
haut = txtVersNum(Tb_prof.Text)
ebret_dess.Profondeur = Round(haut, 2)
If Chk_prof.Value = 0 And Chk_prof.Enabled Then
    Chk_prof.Value = 1
End If
bkprof = True
Cmd_schema.Enabled = False
If chang_prof Then
    chang_prof = False
    If Me.Chk_larg.Value + Me.Chk_long.Value + Me.Chk_prof.Value + Me.Chk_rap.Value = 2 Then
        Call calcul_dimension
    End If
'    Call dess_stock
End If
  End If
End If

 sval_champ = ""
 bkprof = False
 bKP = False
End Sub
Private Sub Tb_prof_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    chang_prof = True
    bKP = True
    sval_champ = Tb_prof.Text
    iSels = Tb_prof.SelStart
    iSell = Tb_prof.SelLength
'    If Len(Tb_prof.Text) <= Tb_prof.MaxLength Then
'        KeyAscii = verif_car(Tb_prof.Text, KeyAscii, "Saisie de la hauteur d'eau", "R")
'    End If
End If
End Sub

Private Sub Tb_Qf_Change()
 Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_Qf.Text, "Saisie du débit de fuite de la retenue", "R")
  If nom = "" Then
    Tb_Qf.Text = sval_champ
    Tb_Qf.SelStart = iSels
    Tb_Qf.SelLength = iSell
  End If
End If
'****

   ebret.qf = txtVersNum(Me.Tb_Qf.Text)
 '   Call ini_ebret_dess
    Call reini_valeurs
     sval_champ = ""
    bKP = False
End Sub
Private Sub Tb_Qf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Qf.Text
    iSels = Tb_Qf.SelStart
    iSell = Tb_Qf.SelLength
'    If Len(Tb_Qf.Text) <= Tb_Qf.MaxLength Then
'        KeyAscii = verif_car(Tb_Qf.Text, KeyAscii, "Saisie du débit de fuite de la retenue", "I")
'    End If
End If
End Sub
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, "" 'Dimensionnement d'un bassin de rétention"
End Sub


Private Sub Tb_rap_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String
bkrap = False
If bKP Then
        nom = verif_cart0(Tb_rap.Text, "Saisie du rapport (largeur/hauteur d'eau)", "R")
  If nom = "" Then
    Tb_rap.Text = sval_champ
    Tb_rap.SelStart = iSels
    Tb_rap.SelLength = iSell
  Else
'  End If
'End If
'****
bkrap = True
rap = txtVersNum(Tb_rap.Text)
ebret_dess.Rapport = Round(rap, 2)
If Chk_rap.Value = 0 And Chk_rap.Enabled Then
    Chk_rap.Value = 1
End If
bkrap = True
Cmd_schema.Enabled = False
If chang_rap Then
    chang_rap = False
    If Me.Chk_larg.Value + Me.Chk_long.Value + Me.Chk_prof.Value + Me.Chk_rap.Value = 2 Then
        Call calcul_dimension
    End If
'    Call dess_stock
End If
  End If
End If
 sval_champ = ""
 bKP = False
bkrap = False

End Sub
Private Sub Tb_rap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    chang_rap = True
     bKP = True
     sval_champ = Tb_rap.Text
    iSels = Tb_rap.SelStart
    iSell = Tb_rap.SelLength
'   If Len(Tb_rap.Text) <= Tb_rap.MaxLength Then
'        KeyAscii = verif_car(Tb_rap.Text, KeyAscii, "Saisie du rapport (largeur/hauteur d'eau)", "R")
'    End If
End If
End Sub

Private Sub Tb_titre_Change()
   Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub

