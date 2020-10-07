VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_stock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Bassin de  Stockage"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "Frm_stock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.TextBox Tb_volume 
      BackColor       =   &H80000016&
      Height          =   405
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Frame Frm_type 
      Caption         =   "Type de bassin"
      Height          =   615
      Left            =   5280
      TabIndex        =   27
      Top             =   1920
      Width           =   4335
      Begin VB.OptionButton Opt_cond 
         Caption         =   "conduite"
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Type de bassin"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Opt_rect 
         Caption         =   "rectangulaire"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Type de bassin"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Opt_cir 
         Caption         =   "circulaire"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Type de bassin"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frm_rect 
      Height          =   1665
      Left            =   5280
      TabIndex        =   38
      Top             =   2400
      Width           =   4335
      Begin VB.CommandButton Cmd_resu 
         Caption         =   "Dessiner"
         Height          =   255
         Left            =   3240
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Dimensionnement et dessin du bassin de stockage"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CheckBox Chk_rap 
         Height          =   200
         Left            =   120
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1220
         Width           =   200
      End
      Begin VB.CheckBox Chk_prof 
         Height          =   200
         Left            =   120
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   900
         Width           =   200
      End
      Begin VB.CheckBox Chk_larg 
         Height          =   200
         Left            =   120
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   600
         Width           =   200
      End
      Begin VB.CheckBox Chk_long 
         Height          =   200
         Left            =   120
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   290
         Width           =   200
      End
      Begin VB.TextBox Tb_rap 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   46
         Top             =   1180
         Width           =   900
      End
      Begin VB.TextBox Tb_prof 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   45
         Top             =   860
         Width           =   900
      End
      Begin VB.TextBox Tb_larg 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   44
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox Tb_long 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   43
         Top             =   220
         Width           =   900
      End
      Begin VB.Label Lb_urap 
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         Top             =   1365
         Width           =   60
      End
      Begin VB.Label Lb_uprof 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Lb_ularg 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   48
         Top             =   580
         Width           =   300
      End
      Begin VB.Label Lb_ulong 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Lb_intrap 
         Caption         =   "Rapport l/h"
         Height          =   255
         Left            =   580
         TabIndex        =   42
         Top             =   1220
         Width           =   1215
      End
      Begin VB.Label Lb_intprof 
         Caption         =   "Hauteur d'eau"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Lb_intlarg 
         Caption         =   "Largeur"
         Height          =   255
         Left            =   585
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lb_intlong 
         Caption         =   "Longueur"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frm_circ 
      Height          =   1455
      Left            =   5280
      TabIndex        =   31
      Top             =   2400
      Width           =   4335
      Begin VB.TextBox Tb_haut 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   33
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox Tb_diam 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   32
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Lb_inthaut 
         Caption         =   "Hauteur d'eau"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Lb_intdiam 
         Caption         =   "Diamètre"
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   770
         Width           =   1215
      End
      Begin VB.Label Lb_uhaut 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label Lb_udiam 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   765
         Width           =   375
      End
   End
   Begin VB.ComboBox Cb_stockage 
      Height          =   315
      Left            =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   4000
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du volume du bassin de stockage"
      Top             =   3680
      Width           =   1000
   End
   Begin VB.TextBox Tb_Qav 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "0"
      Top             =   3650
      Width           =   900
   End
   Begin VB.Frame Frm_bassin 
      Caption         =   "Bassin "
      Height          =   3600
      Left            =   240
      TabIndex        =   0
      Top             =   5
      Width           =   4695
      Begin VB.CommandButton Cmd_Sel_Bv 
         Caption         =   "Sélection d'un bassin versant"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3180
         Width           =   4215
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   2925
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2700
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   2925
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2325
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   2925
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1980
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   2925
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1620
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2925
         MaxLength       =   8
         TabIndex        =   1
         Top             =   285
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   2925
         MaxLength       =   8
         TabIndex        =   2
         Top             =   645
         Width           =   900
      End
      Begin VB.TextBox Tb_bv 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   2925
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Lb_ubv 
         Height          =   255
         Index           =   5
         Left            =   4005
         TabIndex        =   26
         Top             =   2370
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "mn"
         Height          =   255
         Index           =   6
         Left            =   4005
         TabIndex        =   23
         Top             =   2745
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "ha"
         Height          =   255
         Index           =   4
         Left            =   4005
         TabIndex        =   22
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "l/ha/s"
         Height          =   255
         Index           =   3
         Left            =   4005
         TabIndex        =   21
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "l/s"
         Height          =   255
         Index           =   2
         Left            =   4005
         TabIndex        =   20
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "l/s"
         Height          =   255
         Index           =   1
         Left            =   4005
         TabIndex        =   19
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Lb_ubv 
         Caption         =   "l/s"
         Height          =   255
         Index           =   0
         Left            =   4005
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Temps de concentration du B.V."
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2805
         Width           =   2535
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Coefficient de ruissellement du B.V."
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2445
         Width           =   2535
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Surface du B.V."
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   2085
         Width           =   2535
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Intensité de pluie de référence"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1665
         Width           =   2415
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Débit d'eau pluviale "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Débit de temps sec "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label Lb_intbv 
         Caption         =   "Débit de référence"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1050
         Width           =   2295
      End
   End
   Begin VB.Frame Frm_cond 
      Height          =   1500
      Left            =   5280
      TabIndex        =   57
      Top             =   2520
      Width           =   4335
      Begin VB.TextBox Tb_longc 
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   59
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Tb_diamc 
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   58
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Lb_ulongc 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   63
         Top             =   770
         Width           =   375
      End
      Begin VB.Label Lb_udiamc 
         Caption         =   "m"
         Height          =   255
         Left            =   3000
         TabIndex        =   62
         Top             =   410
         Width           =   375
      End
      Begin VB.Label Lb_intlongc 
         Caption         =   "Longueur"
         Height          =   255
         Left            =   600
         TabIndex        =   61
         Top             =   770
         Width           =   1215
      End
      Begin VB.Label Lb_intdiamc 
         Caption         =   "Diamètre"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   410
         Width           =   1215
      End
   End
   Begin RichTextLib.RichTextBox tb_resu 
      Height          =   1350
      Left            =   5280
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2381
      _Version        =   393217
      BackColor       =   -2147483626
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Frm_stock.frx":08CA
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   5880
      MaxLength       =   30
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Lb_uqav 
      Caption         =   "l/s"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   3700
      Width           =   255
   End
   Begin VB.Label Lb_Qav 
      Caption         =   "Débit aval admissible"
      Height          =   345
      Left            =   240
      TabIndex        =   17
      Top             =   3700
      Width           =   1575
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Bassin de stockage"
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
Attribute VB_Name = "Frm_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private okg As Boolean
Private owner As MDIFrm_menu
Private esave As st_savstock
Public nom_ouvrage As String
'Private nom_fich As String
Public nom_type As String
Private lhFicDbf As Long
Private FileLength As Integer
Private chang_diam As Boolean
Private chang_haut As Boolean
Private chang_diamc As Boolean
Private chang_longc As Boolean
Private chang_long As Boolean
Private chang_larg As Boolean
Private chang_prof As Boolean
Private chang_rap As Boolean
Private chang_chk As Boolean
Private nombassin As String
Private list_don1() As Variant
Private list_int1() As Variant
Private list_resu1() As Variant
Private ebstock_dess As stock_dess
Private st_texte As String
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
'    Case Is = "Tb_Qav"
'         nom1 = "Lb_Qav"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intbv"
'         Lb_intbv(index_prec).ForeColor = coulp
'    Case Is = "Lb_Qav"
'         Lb_Qav.ForeColor = coulp
'    Case Is = "Frm_bassin"
'         Frm_bassin.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intbv"
'         Lb_intbv(Index).ForeColor = coul
'    Case Is = "Lb_Qav"
'         Lb_Qav.ForeColor = coul
'    Case Is = "Frm_bassin"
'         Frm_bassin.ForeColor = coul
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
    Case Is = "Lb_Qav"
         Tb_Qav.SetFocus
    Case Is = "Frm_bassin"
         Tb_bv(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_intbv", "Tb_bv", "Frm_bassin"
        mes = IDhlp_StockageOrigineMethode '"Origine de la méthode"
    Case Is = "Cmd_calcul" '"Lb_Qav", "Tb_Qav"
        mes = IDhlp_StockagePresentationMethodeCalcul '"Présentation de la méthode de calcul"
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
l1 = Array(0, "TB_bv", "TB_qav", "CMD_calcul", "Tb_long", "Tb_larg", "Tb_prof", "Tb_rap", "Tb_diam", "Tb_haut", "Tb_diamc", "Tb_longc")
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
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc) ' 10040 '200
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub

Public Sub dess_stock()
Dim ebvolume As volume_dess
ebvolume.coef = ebstock_dess.coef
ebvolume.Diametre = ebstock_dess.Diametre
ebvolume.hauteur = ebstock_dess.hauteur
ebvolume.Diametrec = ebstock_dess.Diametrec
ebvolume.Longueurc = ebstock_dess.Longueurc
ebvolume.Largeur = ebstock_dess.Largeur
ebvolume.Longueur = ebstock_dess.Longueur
ebvolume.Profondeur = ebstock_dess.Profondeur
ebvolume.Rapport = ebstock_dess.Rapport
owner.fdessin.UC_graphique1.graphique_clear
Frm_desprint.UC_graphique1.graphique_clear
' impression false
            Me.mnuprint.Enabled = False
Select Case ebstock_dess.type
    Case Is = "rect"
        If ebstock_dess.Longueur > 0 And ebstock_dess.Largeur > 0 And ebstock_dess.Profondeur > 0 Then
            Call init_graph_rect(owner.fdessin.UC_graphique1, ebvolume)
            Call dess_stock_rect(owner.fdessin.UC_graphique1, ebvolume)
            Call init_graph_rect(Frm_desprint.UC_graphique1, ebvolume)
            Call dess_stock_rect(Frm_desprint.UC_graphique1, ebvolume)
' impression true
            Me.mnuprint.Enabled = True
        End If
    Case Is = "circ"
        If ebstock_dess.hauteur > 0 And ebstock_dess.Diametre > 0 Then
            Call init_graph_circ(owner.fdessin.UC_graphique1, ebvolume)
            Call dess_stock_circ(owner.fdessin.UC_graphique1, ebvolume)
            Call init_graph_circ(Frm_desprint.UC_graphique1, ebvolume)
            Call dess_stock_circ(Frm_desprint.UC_graphique1, ebvolume)
' impression true
            Me.mnuprint.Enabled = True
        End If
    Case Is = "cond"
        If ebstock_dess.Longueurc > 0 And ebstock_dess.Diametrec > 0 Then
           Call init_graph_cond(owner.fdessin.UC_graphique1, ebvolume)
           Call dess_stock_cond(owner.fdessin.UC_graphique1, ebvolume)
           Call init_graph_cond(Frm_desprint.UC_graphique1, ebvolume)
           Call dess_stock_cond(Frm_desprint.UC_graphique1, ebvolume)
' impression true
            Me.mnuprint.Enabled = True
        End If
End Select
    ouv_sauve = True
End Sub
Private Sub lect_fich()
Dim za As st_savstock
Dim za1 As st_savsto1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_stockage.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavstock
        If Trim(za.type) = nom_type Then
            Cb_stockage.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
st_texte = Cb_stockage.list(0)
Cb_stockage.Text = Cb_stockage.list(0)
Cb_stockage.Refresh
 
Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub

Private Sub Cb_stockage_Change()
    Cb_stockage.Text = st_texte
End Sub

Public Sub Cb_stockage_click()
Dim za As st_savstock
Dim za1 As st_savsto1
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    st_texte = Trim(nom_ouvrage)
    Cb_stockage.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavstock
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_stockage.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre
            ebstock = za.stockage
            ebstock_dess = ebstock.dessstock
            nombassin = ebstock.nombv
            Call ini_form
'           Call reini_valeurs
            Call ini_lbresu
'           Me.Cmd_del.Visible = True
            If Trim(nombassin) <> "" Then
                ebv.Qchoisi = ""
                Close #lhFicDbf

                Call rec_bassin(nombassin, "versant")
                
                If Trim(ebv.Qchoisi) <> "" Then
                    Me.Frm_bassin.Caption = "Bassin versant : " + Trim(nombassin) 'Trim(ebv.nom)
                    Select Case ebv.Qchoisi
                        Case Is = "CAQUOT"
                            Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (CAQUOT)"
                        Case Is = "RATION"
                            Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (Rationnelle)"
                        Case Is = "HYDROG"
                            Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (Hydrogramme)"
                    End Select
                Else
                    Me.Frm_bassin.Caption = "Bassin versant : " + Trim(nombassin)
                    Me.Lb_intbv(0).Caption = "Débit d'eau pluviale "
                End If
            Else
                Me.Frm_bassin.Caption = "Bassin versant : "
                Me.Lb_intbv(0).Caption = "Débit d'eau pluviale "
            End If
                Call reini_valeurs
            If Cmd_calcul.Enabled Then
                Call Cmd_calcul_Click 'Calc_volume
'            Else
'                Call reini_valeurs
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


Private Sub Cb_stockage_KeyDown(KeyCode As Integer, Shift As Integer)
    st_texte = Cb_stockage.Text
    Cb_stockage.Text = st_texte
End Sub

Private Sub Cb_stockage_KeyPress(KeyAscii As Integer)
    st_texte = Cb_stockage.Text
End Sub

Private Sub Chk_larg_Click()
If Me.Chk_larg.Enabled Then
    Call check_volume
End If
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
Call calcul_dim1
Else
 Call check_volume_enable(True)
'  Call check_volume_saisie
    ebstock_dess.opt_larg = (Me.Chk_larg.Value = 1)
    ebstock_dess.opt_long = (Me.Chk_long.Value = 1)
    ebstock_dess.opt_prof = (Me.Chk_prof.Value = 1)
    ebstock_dess.opt_rap = (Me.Chk_rap.Value = 1)
End If
End Sub
Private Sub check_volume_saisie()

    Me.Tb_larg.Enabled = Me.Chk_larg.Enabled
    Me.Tb_long.Enabled = Me.Chk_long.Enabled
    Me.Tb_prof.Enabled = Me.Chk_prof.Enabled
    Me.Tb_rap.Enabled = Me.Chk_rap.Enabled
'    ebstock_dess.opt_larg = (Me.Chk_larg.Value = 1)
'    ebstock_dess.opt_long = (Me.Chk_long.Value = 1)
'    ebstock_dess.opt_prof = (Me.Chk_prof.Value = 1)
'    ebstock_dess.opt_rap = (Me.Chk_rap.Value = 1)
End Sub
Private Sub check_volume_recup()
        If ebstock_dess.opt_long Then
        Me.Chk_long.Value = 1
        Else
        Me.Chk_long.Value = 0
    End If
    If ebstock_dess.opt_larg Then
        Me.Chk_larg.Value = 1
         Else
        Me.Chk_larg.Value = 0
   End If
    If ebstock_dess.opt_prof Then
        Me.Chk_prof.Value = 1
        Else
        Me.Chk_prof.Value = 0
    End If
    If ebstock_dess.opt_rap Then
        Me.Chk_rap.Value = 1
    Else
        Me.Chk_rap.Value = 0
    End If
    Me.Chk_larg.Enabled = ebstock_dess.opt_larg
    Me.Chk_long.Enabled = ebstock_dess.opt_long
    Me.Chk_prof.Enabled = ebstock_dess.opt_prof
    Me.Chk_rap.Enabled = ebstock_dess.opt_rap

    Me.Tb_larg.Enabled = Me.Chk_larg.Enabled
    Me.Tb_long.Enabled = Me.Chk_long.Enabled
    Me.Tb_prof.Enabled = Me.Chk_prof.Enabled
    Me.Tb_rap.Enabled = Me.Chk_rap.Enabled
End Sub

Private Sub check_volume_enable(ByVal ok As Boolean)
  Me.Chk_larg.Enabled = ok
  Me.Chk_long.Enabled = ok
  Me.Chk_prof.Enabled = ok
  Me.Chk_rap.Enabled = ok
  Call check_volume_saisie
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

Private Sub Cmd_calcul_Click()
     Dim mes As String
    Dim nom As String
    Dim reponse As Integer
    nom = "Cmd_calcul"
    mes = Rec_Mes(nom, 0)
    Change_Focus nom, 0
    owner.affich_aide Me.Name, mes
   Call key13(Me)
    Call Calc_volume
    ebstock.dessstock = ebstock_dess
    
    Call init_form_dess
'    Call reini_dess
End Sub
Private Sub reini_dess()

    Me.Frm_type.Visible = True
    Me.Opt_cir.Value = False
    Me.Opt_rect.Value = False
    Me.Opt_cond.Value = False
'    Me.Tb_diam.Text = ""
'    Me.Tb_haut.Text = ""
'    Me.Tb_diamc.Text = ""
'    Me.Tb_longc.Text = ""
'    Me.Tb_long.Text = ""
'    Me.Tb_larg.Text = ""
'    Me.Tb_prof.Text = ""
'    Me.Tb_rap.Text = ""
    Call check_volume_recup
    Call init_form_dess
'    chang_diam = True
'    Me.Tb_diam.Text = Format(ebstock.dessstock.Diametre, "####0.00")
'    Me.Tb_long.Text = Format(ebstock.dessstock.Longueur, "####0.00")
'    Me.Tb_larg.Text = Format(ebstock.dessstock.Largeur, "####0.00")
'    Me.Tb_prof.Text = Format(ebstock.dessstock.Profondeur, "####0.00")
'    Me.Tb_rap.Text = Format(ebstock.dessstock.Rapport, "####0.00")

End Sub


Private Sub meAffiche()
    DoEvents
    Me.Show
End Sub
Private Sub Calc_volume()
Dim Q As Double, a As Double, Qpav As Double, Vr As Double, v As Double, Ipcav As Double
Dim sresult As String, sresult1 As String
'calcul de Qpav
    Call modi_lbresu
    Qpav = ebstock.Qav - ebstock.Qts
    sresult = " Calcul du volume maximum stocké "
'calcul de Ipcav
    Ipcav = Qpav / (ebstock.imper / 100# * ebstock.surface)
    sresult = sresult + Chr(13) + Chr(10) + "  Intensité de pluie aval   = " + ajout_zero(Trim(str(Round(Ipcav, 3)))) + " l/ha/s"
    ebstock.Ipcav = Round(Ipcav, 3)
a = recup_alphat(ebstock.tc)
' recherche du volume réduit
      Vr = rec_vr(Ipcav, ebstock.lcrin)
  '  Vr = 100#
    sresult = sresult + Chr(13) + Chr(10) + "  Volume réduit   = " + ajout_zero(Trim(str(Round(Vr, 3)))) + " m3/ha"
    ebstock.Vr = Round(Vr, 3)
'calcul du volume
    v = Vr * ebstock.imper / 100# * ebstock.surface * a
    sresult = sresult + Chr(13) + Chr(10) + "  Facteur lié au temps de concentration    = " + ajout_zero(Trim(str(Round(a, 3))))
    ebstock.alphat = Round(a, 3)
    sresult1 = "  Volume de stockage   = " + ajout_zero(Trim(str(Round(v, 3)))) + " m3"
    ebstock.volume = Round(v, 3)
    ouv_sauve = True
    Me.tb_resu.Text = sresult
    Me.Tb_volume.Text = sresult1
    Me.Cmd_calcul.Enabled = False
End Sub
Private Function recup_alphat(ByVal tce As Double) As Double
Dim list_TC(13, 2) As Double
Dim a As Double

Dim i As Integer

list_TC(1, 1) = 10
list_TC(2, 1) = 15
list_TC(3, 1) = 20
list_TC(4, 1) = 25
list_TC(5, 1) = 30
list_TC(6, 1) = 35
list_TC(7, 1) = 40
list_TC(8, 1) = 50
list_TC(9, 1) = 60
list_TC(10, 1) = 80
list_TC(11, 1) = 100
list_TC(12, 1) = 120
list_TC(13, 1) = 180
list_TC(1, 2) = 1.25
list_TC(2, 2) = 1.48
list_TC(3, 2) = 1.63
list_TC(4, 2) = 1.74
list_TC(5, 2) = 1.82
list_TC(6, 2) = 1.88
list_TC(7, 2) = 1.93
list_TC(8, 2) = 2.02
list_TC(9, 2) = 2.06
list_TC(10, 2) = 2.12
list_TC(11, 2) = 2.17
list_TC(12, 2) = 2.2
list_TC(13, 2) = 2.25



i = 1
While tce > list_TC(i, 1) And i < UBound(list_TC)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
a = (tce - list_TC(i - 1, 1)) * (list_TC(i, 2) - list_TC(i - 1, 2)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 2)
recup_alphat = a
End Function



Private Sub Cmd_resu_Click()
'    ebstock_dess.Longueur = txtVersNum(Tb_long.Text)
'    ebstock_dess.largeur = txtVersNum(Tb_larg.Text)
'    ebstock_dess.profondeur = txtVersNum(Tb_prof.Text)
'    owner.fdessin.UC_graphique1.graphique_clear
'    Call init_graph_rect(owner.fdessin.UC_graphique1)
'    Call dess_stock_rect(owner.fdessin.UC_graphique1, ebstock_dess)

End Sub

Private Sub Cmd_Sel_Bv_Click()
    Dim pict1 As New StdPicture
    dess_anc = chemin_app + "dessanc.bmp"
    If Dir(dess_anc) <> "" Then
        Kill dess_anc
    End If
    Set pict1 = owner.fdessin.UC_graphique1.lire_pict1()
    SavePicture pict1, chemin_app + "dessanc.bmp"
    Me.Enabled = False
    sto_bv = True
    owner.fdessin.UC_graphique1.graphique_clear
    Set owner.fbassin = New Frm_bv2
    owner.fbassin.Show
    owner.fbassin.nom_ouvrage = nombassin
    owner.fbassin.Cmd_retour.Visible = True
    owner.fbassin.Cmd_retour.Caption = "Retour au bassin de stockage"
    fich_lect = nom_fich
    Call owner.fbassin.rec_bassin_versant
    owner.affich_aide owner.fbassin.Name, "Module" '"Calcul de débit de bassin versant"
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Public Sub reini_valeurs()
 '   me.Tb_Qav.Text = "0"
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.init_titleb ""
    Call ini_lbresu
            ' impression false
                    Me.mnuprint.Enabled = False

'And ebstock.Qrin > 0 And ebstock.Qts > 0
    If ebstock.Qpluie > 0 And ebstock.lcrin > 0 And ebstock.surface > 0 _
    And ebstock.imper > 0 And ebstock.tc > 0 And ebstock.Qrin > 0 _
    And ebstock.Qts > 0 And ebstock.Qav > 0 Then
        Me.Cmd_calcul.Enabled = True
    Else
        Me.Cmd_calcul.Enabled = False
    End If
        Me.Frm_type.Visible = False
        Me.Frm_circ.Visible = False
        Me.Frm_rect.Visible = False
        Me.Frm_cond.Visible = False
    ouv_sauve = True
End Sub
Private Sub ini_form()
    Me.Tb_bv(0).Text = rempl_virgule(Format(ebstock.Qpluie, "####0.0"))  '* 1000
    Me.Tb_bv(2).Text = rempl_virgule(Format(ebstock.Qrin, "####0.0"))
    Me.Tb_bv(1).Text = rempl_virgule(Format(ebstock.Qts, "####0.0"))
    Me.Tb_bv(3).Text = rempl_virgule(Format(ebstock.lcrin, "###0"))
'Houpie 2005/03/21
'    Me.Tb_bv(4).Text = rempl_virgule(Format(ebstock.surface, "###0"))
     Me.Tb_bv(4).Text = rempl_virgule(Format(ebstock.surface, "###0.00"))
    Me.Tb_bv(5).Text = rempl_virgule(Format(ebstock.imper, "###0"))
    Me.Tb_bv(6).Text = rempl_virgule(Format(ebstock.tc, "###0.0"))
    Me.Tb_Qav.Text = rempl_virgule(Format(ebstock.Qav, "###0"))
    ebstock_dess = ebstock.dessstock
    Call init_form_dess
  '  ebstock_dess = ebstock.dessstock
End Sub
Private Sub init_form_dess()
Call check_volume_enable(False)
Chk_larg.Value = 0
Chk_long.Value = 0
Chk_prof.Value = 0
Chk_rap.Value = 0
'    ebstock_dess = ebstock.dessstock
    Me.Tb_haut.Text = rempl_virgule(Format(ebstock_dess.hauteur, "####0.00"))
    chang_diam = True
    Me.Tb_diam.Text = rempl_virgule(Format(ebstock_dess.Diametre, "####0.00"))
    chang_diam = False
    Me.Tb_diamc.Text = rempl_virgule(Format(ebstock_dess.Diametrec, "####0.00"))
    Me.Tb_longc.Text = rempl_virgule(Format(ebstock_dess.Longueurc, "####0.00"))
    Me.Tb_long.Text = rempl_virgule(Format(ebstock_dess.Longueur, "####0.00"))
    Me.Tb_larg.Text = rempl_virgule(Format(ebstock_dess.Largeur, "####0.00"))
    Me.Tb_prof.Text = rempl_virgule(Format(ebstock_dess.Profondeur, "####0.00"))
    Me.Tb_rap.Text = rempl_virgule(Format(ebstock_dess.Rapport, "####0.00"))
    Call check_volume_recup
    Call check_volume
'    Call check_volume_enable(True)
 Call calcul_dim1
  
    If Trim(ebstock_dess.type) <> "" Then
        Me.Frm_type.Visible = True
        Select Case ebstock_dess.type
        Case Is = "circ"
            Opt_cir.Value = True
            Opt_rect.Value = False
            Opt_cond.Value = False
           Call Opt_cir_Click

        Case Is = "rect"
            Opt_cir.Value = False
            Opt_rect.Value = True
            Opt_cond.Value = False
            Call Opt_rect_Click
        Case Is = "cond"
            Opt_cir.Value = False
            Opt_rect.Value = False
            Opt_cond.Value = True
            Call Opt_cond_Click
        End Select
        Call dess_stock
    Else
        Opt_cir.Value = False
        Opt_rect.Value = False
        Opt_cond.Value = False
         Me.Frm_type.Visible = True
   End If

End Sub
Public Sub ini_debit(ByVal nom As String)
    nombassin = nom
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
'    owner.fdessin.UC_graphique1.Visible = False
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
    If Trim(ebv.Qchoisi) <> "" Then
        Me.Frm_bassin.Caption = "Bassin versant : " + Trim(nombassin) 'Trim(ebv.nom)
        Select Case ebv.Qchoisi
            Case Is = "CAQUOT"
                Me.Tb_bv(0).Text = rempl_virgule(Format(ebv.Qcor * 1000, "####0.0"))
                Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (CAQUOT)"
            Case Is = "RATION"
                Me.Tb_bv(0).Text = rempl_virgule(Format(ebv.Qmr * 1000, "####0.0"))
                Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (Rationnelle)"
            Case Is = "HYDROG"
                Me.Tb_bv(0).Text = rempl_virgule(Format(ebv.Qhydro * 1000, "####0.0"))
                Me.Lb_intbv(0).Caption = "Débit d'eau pluviale (Hydrogramme)"
        End Select
'julienne 2001/12/12
'        Me.Tb_titre.Text = ""
'        Me.Cb_stockage.Text = ""
        Me.Tb_bv(2).Text = rempl_virgule(Format(ebv.Qrin, "###0.0"))
        Me.Tb_bv(1).Text = rempl_virgule(Format(ebv.Qts, "###0.0"))
        Me.Tb_bv(3).Text = rempl_virgule(Format(eph.lcrin, "###0"))
'Houpie 2005/03/21
        Me.Tb_bv(4).Text = rempl_virgule(Format(ebv.surface, "###0.00"))
'        Me.Tb_bv(4).Text = ajout_zero(Trim(Str(ebv.surface)))
        Me.Tb_bv(5).Text = rempl_virgule(Format(ebv.imper, "###0"))
        Me.Tb_bv(6).Text = rempl_virgule(Format(ebv.tc, "###0.0"))
'        Me.Tb_Qav = 0
'        Frm_do.UC_graphique1.ecr_texta 1665, 1140, "Surface = " + Str(ebv.surface) + " Ha", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 1620, 1530, "Coef. de ruissellement = " + Str(ebv.imper), "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2055, 1995, "Nombre d'habitants = " + Str(ebv.nhab), "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2160, 2505, "Taux de dilution = " + Str(ebv.tdilu), "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2875, 540, "Longueur = " + Str(ebv.lghydr) + " m", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 3145, 810, "Pente = " + Str(ebv.phydr) + " (1/10000)", "G", "B"
'        Frm_do.SSTab1.TabEnabled(1) = True
    Else
        Me.Frm_bassin.Caption = "Bassin versant : "
        Me.Tb_bv(0).Text = "0.0"
        Me.Tb_bv(2).Text = "0.0"
        Me.Tb_bv(1).Text = "0.0"
        Me.Tb_bv(3).Text = "0"
'Houpie 2005/03/21
'        Me.Tb_bv(4).Text = "0"
        Me.Tb_bv(4).Text = "0.00"
        Me.Tb_bv(5).Text = "0"
        Me.Tb_bv(6).Text = "0.0"

'        Frm_do.UC_graphique1.ecr_texta 1665, 1140, "Surface", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 2875, 540, "Longueur", "G", "B"
'        Frm_do.UC_graphique1.ecr_texta 3145, 810, "Pente", "G", "B"
'        Frm_do.SSTab1.TabEnabled(1) = False
   End If
'        Call ini_ebstock_dess
'        Call init_form_dess
        Call reini_valeurs
'        If Cmd_calcul.Enabled Then
'            Call Calc_volume
'            ebstock.dessstock = ebstock_dess
'            Call init_form_dess
'        End If
End Sub
Private Sub ini_lbresu()
 '   Me.tb_resu.BackColor = &H8000000B
    Me.tb_resu.BorderStyle = 1
    Me.tb_resu.Text = ""
'    Me.Tb_volume.BackColor = &H8000000B
    Me.Tb_volume.BorderStyle = 1
    Me.Tb_volume.Text = ""
End Sub
Private Sub modi_lbresu()
'    Me.Tb_volume.BackColor = &H80000009
    Me.Tb_volume.BorderStyle = 1
'    Me.tb_resu.BackColor = &H80000009
    Me.tb_resu.BorderStyle = 1
End Sub
Public Sub ini_ebstock_dess()
    ebstock_dess.type = " "
    ebstock_dess.coef = 0
    ebstock_dess.Diametre = 0#
    ebstock_dess.hauteur = 0#
    ebstock_dess.Diametrec = 0#
    ebstock_dess.Longueurc = 0#
    ebstock_dess.Longueur = 0#
    ebstock_dess.Largeur = 0#
    ebstock_dess.Profondeur = 0#
    ebstock_dess.Rapport = 0#
    ebstock_dess.opt_long = False
    ebstock_dess.opt_larg = False
    ebstock_dess.opt_prof = False
    ebstock_dess.opt_rap = False
End Sub
Public Sub ini_ebstock()
    ebstock.nom = ""
    ebstock.Qpluie = 0#
    ebstock.Qts = 0#
    ebstock.Qrin = 0#
    ebstock.lcrin = 0
    ebstock.surface = 0
    ebstock.imper = 0
    ebstock.tc = 0#
    ebstock.Qav = 0#
    ebstock.Ipcav = 0#
    ebstock.Vr = 0#
    ebstock.alphat = 0#
    ebstock.volume = 0#
    Call ini_ebstock_dess
    ebstock.dessstock = ebstock_dess
End Sub
Private Sub m_quitter_Click()
    Unload owner
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
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


Private Sub Lb_intbv_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intbv"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Qav_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Qav"
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
    reponse = MsgBox("Le bassin de stockage  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de stockage")
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
    reponse = MsgBox("Le bassin de stockage  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de stockage")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_stockage.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de stockage "
            frmf.Caption = nom
            frmf.Show (1)
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_stockage_click
            End If
        Case Is = 7
'            Cb_stockage.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin de stockage "
            frmf.Caption = nom
            frmf.Show (1)
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_stockage_click
            End If
    End Select
Else
'    Cb_stockage.Visible = True
    frmf.Label1.Caption = "Recherche d'un bassin de stockage "
    frmf.Caption = nom
    frmf.Show (1)
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_stockage_click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer
ReDim list_don1(Tb_bv.count, 3)
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "stockage"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE BASSIN de STOCKAGE"
FrmPrint.sstitre1 = "Caractéristiques " + Frm_bassin.Caption
FrmPrint.ssTitre2 = "Résultats intermédiaires"
FrmPrint.ssTitre3 = ""
Frm_imp.Type1 = "stockage"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE BASSIN de STOCKAGE"
Frm_imp.sstitre1 = "Caractéristiques " + Frm_bassin.Caption
Frm_imp.ssTitre2 = "Résultats intermédiaires"
Frm_imp.ssTitre3 = ""
For i = 0 To Tb_bv.count - 1
    list_don1(i, 1) = Lb_intbv(i).Caption
    list_don1(i, 2) = Tb_bv(i).Text
    list_don1(i, 3) = Lb_ubv(i).Caption
Next
    list_don1(i, 1) = Lb_Qav.Caption
    list_don1(i, 2) = Tb_Qav.Text
    list_don1(i, 3) = Lb_uqav.Caption

list_int1 = rec_list(tb_resu.Text)
list_resu1 = rec_list(Tb_volume.Text)
list_resu1 = complet_list_resu1(list_resu1)
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
'FrmPrint.Show
Frm_imp.Show 1

End Sub
Private Function complet_list_resu1(ByVal liste1 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
i = -1
Select Case ebstock_dess.type
     Case Is = "circ", "cond"
        ReDim liste(UBound(liste1) + 4, 3)
     Case Is = "rect"
        ReDim liste(UBound(liste1) + 6, 3)
 End Select
For j = 0 To UBound(liste1)
    i = i + 1
'    ReDim Preserve liste(i, 3)
    liste(i, 1) = liste1(j, 1)
    liste(i, 2) = liste1(j, 2)
    liste(i, 3) = liste1(j, 3)
Next
i = i + 1
'ReDim Preserve liste(i, 3)
liste(i, 1) = ""
liste(i, 2) = ""
liste(i, 3) = ""
i = i + 1
'ReDim Preserve liste(i, 3)
liste(i, 1) = "Type de bassin"
liste(i, 3) = ""
Select Case ebstock_dess.type
     Case Is = "circ"
        liste(i, 2) = "circulaire"
        i = i + 1
        liste(i, 1) = Lb_intdiam.Caption
        liste(i, 2) = txtVersNum(Me.Tb_diam.Text)
        liste(i, 3) = Lb_udiam.Caption
        i = i + 1
        liste(i, 1) = Lb_inthaut.Caption
        liste(i, 2) = txtVersNum(Me.Tb_haut.Text)
        liste(i, 3) = Lb_uhaut.Caption
     Case Is = "rect"
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
    Case Is = "cond"
         liste(i, 2) = "conduite"
         i = i + 1
         liste(i, 1) = Lb_intdiamc.Caption
         liste(i, 2) = txtVersNum(Me.Tb_diamc.Text)
         liste(i, 3) = Lb_udiamc.Caption
         i = i + 1
         liste(i, 1) = Lb_intlongc.Caption
         liste(i, 2) = txtVersNum(Me.Tb_longc.Text)
         liste(i, 3) = Lb_ulongc.Caption
End Select
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
Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, "" ' "Dimensionnement d'un bassin de stockage"
    Change_Couleur "Me", 0
End Sub
Private Sub Form_Load()
    okg = True
    Me.KeyPreview = True
    Call ini_tooltip_stock(Me)
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
''''''    owner.affich_aide Me.Name, "Stockage"
'    nom_fich = chemin_app + "stockage.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "stockage"
    fen_titre = Me.Caption
    ouv_sauve = False
    save_fich = False
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Cb_stockage.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
    nombassin = ""
    Call debut
End Sub
Private Sub debut0()
    Cb_stockage.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
    bKP = False
    sval_champ = ""
    Call init_l_tab
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    Me.Frm_bassin.Caption = "Bassin versant : "
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
    Me.Tb_bv(0).Text = "0.00"
    Me.Tb_bv(2).Text = "0.0"
    Me.Tb_bv(1).Text = "0.0"
    Me.Tb_bv(3).Text = "0"
'Houpie 2005/03/21
'    Me.Tb_bv(4).Text = "0"
    Me.Tb_bv(4).Text = "0.00"
    Me.Tb_bv(5).Text = "0"
    Me.Tb_bv(6).Text = "0.0"
    Me.Tb_Qav = 0
    Call ini_ebstock
    Call check_volume_enable(False)
    Me.Tb_diam.Text = rempl_virgule(Format(ebstock_dess.Diametre, "####0.00"))
    Me.Tb_haut.Text = rempl_virgule(Format(ebstock_dess.hauteur, "####0.00"))
    Me.Tb_long.Text = rempl_virgule(Format(ebstock_dess.Longueur, "####0.00"))
    Me.Tb_larg.Text = rempl_virgule(Format(ebstock_dess.Largeur, "####0.00"))
    Me.Tb_prof.Text = rempl_virgule(Format(ebstock_dess.Profondeur, "####0.00"))
    Me.Tb_rap.Text = rempl_virgule(Format(ebstock_dess.Rapport, "####0.00"))
'    Call ini_ebstock
    Call reini_valeurs
    Call check_volume_recup
    If Trim(ebstock_dess.type) <> "" Then
        Me.Frm_type.Visible = True
        Select Case ebstock_dess.type
        Case Is = "circ"
            Call Opt_cir_Click
        Case Is = "rect"
            Call Opt_rect_Click
        Case Is = "cond"
            Call Opt_cond_Click
        End Select
    End If
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin de stockage  n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin de stockage")
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

Private Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If
End Sub

Public Sub save(ByVal bsous As Boolean)
Dim za As st_savstock
Dim za1 As st_savsto1
Dim i As Integer, isave As Integer
Dim reponse As Integer
 
If Trim(Tb_titre.Text) <> "" Then
    Call funlockb
    ebstock.nombv = nombassin
    ebstock.nom = ebv.nom
    ebstock.dessstock = ebstock_dess
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            za = za1.stsavstock
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un bassin de stockage")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "stockage"
            za.nom = Tb_titre.Text
            za.stockage = ebstock
            za1.stsavstock = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        za.type = "stockage"
        za.nom = Tb_titre.Text
        za.stockage = ebstock
        za1.stsavstock = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        st_texte = Trim(Tb_titre.Text)
        Cb_stockage.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom du bassin de stockage n'est pas renseigné.", , "Sauvegarde d'un bassin de stockage")
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
        Frm_titre.Label2.Caption = "Sauvegarde d'un bassin de stockage "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du bassin de stockage " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du bassin de stockage (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub

Private Sub mnusuppr_Click()
Dim za As st_savstock
Dim za1 As st_savsto1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_stockage.Text) <> "" Then
    Call funlockb
     reponse = MsgBox(Trim(Cb_stockage.Text) + " va être supprimé .", 4, "Suppression d'un bassin de stockage")
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
            za = za1.stsavstock
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_stockage.Text)) Then
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
    Call ini_ebstock
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.init_titleb ""
    Me.Frm_bassin.Caption = "Bassin versant : "
    Me.Lb_intbv(0).Caption = "Débit d'eau pluviale "
    Me.Tb_bv(0).Text = "0.0"
    Me.Tb_bv(2).Text = "0.0"
    Me.Tb_bv(1).Text = "0.0"
    Me.Tb_bv(3).Text = "0"
'Houpie 2005/03/21
'    Me.Tb_bv(4).Text = "0"
    Me.Tb_bv(4).Text = "0.00"
    Me.Tb_bv(5).Text = "0"
    Me.Tb_bv(6).Text = "0.0"
    Me.Tb_Qav.Text = 0
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
Private Sub Opt_cir_Click()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    ebstock_dess.type = "circ"
    ebstock_dess.coef = 0.5 'difference entre hauteur et hauteur d'eau
    chang_diam = True
    chang_haut = True
    Me.Frm_circ.Visible = True
    Me.Frm_rect.Visible = False
    Me.Frm_cond.Visible = False
'    If ebstock_dess.hauteur > 0 And ebstock_dess.Diametre > 0 Then
        Call dess_stock
'    End If
End Sub
Private Sub Opt_cond_Click()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    ebstock_dess.type = "cond"
    ebstock_dess.coef = 0.5 'difference entre hauteur et hauteur d'eau
    chang_diamc = True
    chang_longc = True
    Me.Frm_circ.Visible = False
    Me.Frm_rect.Visible = False
    Me.Frm_cond.Visible = True
'    If ebstock_dess.Longueurc > 0 And ebstock_dess.Diametrec > 0 Then
        Call dess_stock
'    End If
End Sub

Private Sub Opt_rect_Click()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    ebstock_dess.type = "rect"
    ebstock_dess.coef = 0.5 'difference entre hauteur et hauteur d'eau
    chang_long = Chk_long.Enabled
    chang_larg = Chk_larg.Enabled
    chang_prof = Chk_prof.Enabled
    chang_rap = Chk_rap.Enabled
    Me.Frm_rect.Visible = True
    Me.Frm_circ.Visible = False
    Me.Frm_cond.Visible = False
'    If ebstock_dess.Longueur > 0 And ebstock_dess.Largeur > 0 And ebstock_dess.Profondeur > 0 Then
        Call dess_stock
'    End If
End Sub

Private Sub Tb_bv_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie débit d'eau pluviale", "R")
            Case Is = 1
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie débit de temps sec", "R")
            Case Is = 2
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie débit de rinçage", "R")
            Case Is = 3
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie de l'intensité de pluie de rinçage", "I")
            Case Is = 4
'Houpie 2005/03/21
'                nom = verif_cart0(Tb_bv(Index).Text, "Saisie de la surface du B.V.", "I")
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie de la surface du B.V.", "R")
            Case Is = 5
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie du coefficient de ruissellement du B.V.", "I")
            Case Is = 6
                nom = verif_cart0(Tb_bv(Index).Text, "Saisie du temps de concentration du B.V.", "R")
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
        ebstock.Qpluie = txtVersNum(Me.Tb_bv(0).Text)
    Case Is = 1
        ebstock.Qts = txtVersNum(Me.Tb_bv(1).Text)
    Case Is = 2
        ebstock.Qrin = txtVersNum(Me.Tb_bv(2).Text)
    Case Is = 3
        ebstock.lcrin = txtVersNum(Me.Tb_bv(3).Text)
    Case Is = 4
        ebstock.surface = txtVersNum(Me.Tb_bv(4).Text)
    Case Is = 5
        ebstock.imper = txtVersNum(Me.Tb_bv(5).Text)
    Case Is = 6
        ebstock.tc = txtVersNum(Me.Tb_bv(6).Text)
End Select
'    Call ini_ebstock_dess
'    ebstock.dessstock = ebstock_dess
'    Call init_form_dess
    Call reini_valeurs
    sval_champ = ""
    bKP = False

End Sub

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

Private Sub Tb_bv_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_bv(Index).Text
    iSels = Tb_bv(Index).SelStart
    iSell = Tb_bv(Index).SelLength
    bKP = True
'    If Len(Tb_bv(Index).Text) <= Tb_bv(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie débit d'eau pluviale", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie débit de temps sec", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie débit de rinçage", "R")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie de l'intensité de pluie de rinçage", "I")
'            Case Is = 4
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie de la surface du B.V.", "I")
'            Case Is = 5
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie du coefficient de ruissellement du B.V.", "I")
'            Case Is = 6
'                KeyAscii = verif_car(Tb_bv(Index).Text, KeyAscii, "Saisie du temps de concentration du B.V.", "R")
'        End Select
'    End If
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

Private Sub Tb_diam_Change()
Dim resu As Double, surf As Double
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_diam.Text, "Saisie du diamètre", "R")
  If nom = "" Then
    Tb_diam.Text = sval_champ
    Tb_diam.SelStart = iSels
    Tb_diam.SelLength = iSell
  Else
'  End If
'End If
'****

If chang_diam Then
    chang_haut = False
    resu = txtVersNum(Tb_diam.Text)
    If resu > 0 Then
        surf = pi * ((resu / 2) ^ 2)
        resu = ebstock.volume / surf
    End If
    Me.Tb_haut.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
'        Me.Tb_haut.Text = Trim(Str(Round(resu, 2)))
    ebstock_dess.hauteur = Round(resu, 2)
    ebstock_dess.Diametre = txtVersNum(Tb_diam.Text)
    Call dess_stock
End If
  End If
End If
 sval_champ = ""
 bKP = False

End Sub

Private Sub Tb_diam_Click()
Call sel_text(Tb_diam)

End Sub

Private Sub Tb_diam_GotFocus()
Call sel_text(Tb_diam)

End Sub

Private Sub Tb_diam_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_diam.Text
    iSels = Tb_diam.SelStart
    iSell = Tb_diam.SelLength
    bKP = True
    chang_diam = True
End Sub

Private Sub Tb_diam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_diam.Text
    iSels = Tb_diam.SelStart
    iSell = Tb_diam.SelLength
    bKP = True
    chang_diam = True
'    If Len(Tb_diam.Text) <= Tb_diam.MaxLength Then
'        KeyAscii = verif_car(Tb_diam.Text, KeyAscii, "Saisie du diamètre", "R")
'    End If
End If
End Sub
Private Sub Tb_diamc_Change()
Dim resu As Double, surf As Double
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_diamc.Text, "Saisie du diamètre", "R")
  If nom = "" Then
    Tb_diamc.Text = sval_champ
    Tb_diamc.SelStart = iSels
    Tb_diamc.SelLength = iSell
  Else
'  End If
'End If
'****

If chang_diamc Then
    chang_longc = False
    resu = txtVersNum(Tb_diamc.Text)
    If resu > 0 Then
        surf = pi * ((resu / 2) ^ 2)
        resu = ebstock.volume / surf
    End If
    Me.Tb_longc.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
'        Me.Tb_haut.Text = Trim(Str(Round(resu, 2)))
    ebstock_dess.Longueurc = Round(resu, 2)
    ebstock_dess.Diametrec = txtVersNum(Tb_diamc.Text)
    Call dess_stock
End If
  End If
End If

 sval_champ = ""
 bKP = False

End Sub

Private Sub Tb_diamc_Click()
Call sel_text(Tb_diamc)

End Sub

Private Sub Tb_diamc_GotFocus()
Call sel_text(Tb_diamc)

End Sub

Private Sub Tb_diamc_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_diamc.Text
    iSels = Tb_diamc.SelStart
    iSell = Tb_diamc.SelLength
    bKP = True
    chang_diamc = True
    chang_diamc = True
End Sub

Private Sub Tb_diamc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_diamc.Text
    iSels = Tb_diamc.SelStart
    iSell = Tb_diamc.SelLength
    bKP = True
    chang_diamc = True
'    If Len(Tb_diamc.Text) <= Tb_diamc.MaxLength Then
'        KeyAscii = verif_car(Tb_diamc.Text, KeyAscii, "Saisie du diamètre", "R")
'    End If
End If
End Sub

Private Sub Tb_haut_Click()
Call sel_text(Tb_haut)

End Sub

Private Sub Tb_haut_GotFocus()
Call sel_text(Tb_haut)

End Sub

Private Sub Tb_haut_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_haut.Text
    iSels = Tb_haut.SelStart
    iSell = Tb_haut.SelLength
    bKP = True
    chang_haut = True
End Sub

Private Sub Tb_larg_Click()
Call sel_text(Tb_larg)

End Sub

Private Sub Tb_larg_GotFocus()
Call sel_text(Tb_larg)

End Sub

Private Sub Tb_larg_KeyDown(KeyCode As Integer, Shift As Integer)
     sval_champ = Tb_larg.Text
    iSels = Tb_larg.SelStart
    iSell = Tb_larg.SelLength
    bKP = True
   chang_larg = True
End Sub

Private Sub Tb_long_Click()
Call sel_text(Tb_long)

End Sub

Private Sub Tb_long_GotFocus()
Call sel_text(Tb_long)

End Sub

Private Sub Tb_long_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_long.Text
    iSels = Tb_long.SelStart
    iSell = Tb_long.SelLength
    bKP = True
    chang_long = True
End Sub

Private Sub Tb_longc_Change()
Dim resu As Double, surf As Double
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_longc.Text, "Saisie de la longueur", "R")
  If nom = "" Then
    Tb_longc.Text = sval_champ
    Tb_longc.SelStart = iSels
    Tb_longc.SelLength = iSell
  Else
'  End If
'End If
'****

If chang_longc Then
    chang_diamc = False
    resu = txtVersNum(Tb_longc.Text)
    If resu > 0 Then
        surf = ebstock.volume / resu
        resu = 2 * Sqr(surf / pi)
    End If
    Me.Tb_diamc.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
'        Me.Tb_diam.Text = Trim(Str(Round(resu, 2)))
'        Lb_resultat = "Diametre = " + Str(Round(resu, 2)) + " m"
    ebstock_dess.Diametrec = Round(resu, 2)
    ebstock_dess.Longueurc = txtVersNum(Tb_longc.Text)
    Call dess_stock
End If
  End If
End If

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_longc_Click()
Call sel_text(Tb_longc)

End Sub

Private Sub Tb_longc_GotFocus()
Call sel_text(Tb_longc)

End Sub

Private Sub Tb_longc_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_longc.Text
    iSels = Tb_longc.SelStart
    iSell = Tb_longc.SelLength
    bKP = True
    chang_longc = True
End Sub

Private Sub Tb_longc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_longc.Text
    iSels = Tb_longc.SelStart
    iSell = Tb_longc.SelLength
    bKP = True
    chang_longc = True
'    If Len(Tb_longc.Text) <= Tb_longc.MaxLength Then
'        KeyAscii = verif_car(Tb_longc.Text, KeyAscii, "Saisie de la longueur", "R")
'    End If
End If
End Sub

Private Sub Tb_haut_Change()
Dim resu As Double, surf As Double
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_haut.Text, "Saisie de la hauteur d'eau", "R")
  If nom = "" Then
    Tb_haut.Text = sval_champ
    Tb_haut.SelStart = iSels
    Tb_haut.SelLength = iSell
  Else
'  End If
'End If
'****

If chang_haut Then
    chang_diam = False
    resu = txtVersNum(Tb_haut.Text)
    If resu > 0 Then
        surf = ebstock.volume / resu
        resu = 2 * Sqr(surf / pi)
    End If
    Me.Tb_diam.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
'        Me.Tb_diam.Text = Trim(Str(Round(resu, 2)))
    Lb_resultat = "Diametre = " + ajout_zero(Trim(str(Round(resu, 2)))) + " m"
    ebstock_dess.Diametre = Round(resu, 2)
    ebstock_dess.hauteur = txtVersNum(Tb_haut.Text)
    Call dess_stock
End If
  End If
End If

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_haut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_haut.Text
    iSels = Tb_haut.SelStart
    iSell = Tb_haut.SelLength
    bKP = True
    chang_haut = True
'    If Len(Tb_haut.Text) <= Tb_haut.MaxLength Then
'        KeyAscii = verif_car(Tb_haut.Text, KeyAscii, "Saisie de la hauteur d'eau", "R")
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
ebstock_dess.Largeur = Round(xlar, 2)
If Chk_larg.Value = 0 And Chk_larg.Enabled Then
    Chk_larg.Value = 1
End If
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
     sval_champ = Tb_larg.Text
    iSels = Tb_larg.SelStart
    iSell = Tb_larg.SelLength
    bKP = True
   chang_larg = True
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
ebstock_dess.Longueur = Round(xlon, 2)
If Chk_long.Value = 0 And Chk_long.Enabled Then
    Chk_long.Value = 1
End If
    bklong = True
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
Private Sub calcul_dim1()
Dim resu As Double, surf As Double
    resu = txtVersNum(Tb_diam.Text)
    If resu > 0 Then
        surf = pi * ((resu / 2) ^ 2)
        resu = ebstock.volume / surf
    End If
    Me.Tb_haut.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
    ebstock_dess.hauteur = Round(resu, 2)
    resu = txtVersNum(Tb_diamc.Text)
    If resu > 0 Then
        surf = pi * ((resu / 2) ^ 2)
        resu = ebstock.volume / surf
    End If
    Me.Tb_longc.Text = rempl_virgule(Format(Round(resu, 2), "##0.00"))
    ebstock_dess.Longueurc = Round(resu, 2)

End Sub
Private Sub calcul_dimension()
Dim surf As Double, xlon As Double, xlar As Double, haut As Double, rap As Double
    ebstock_dess.opt_larg = (Me.Chk_larg.Value = 1)
    ebstock_dess.opt_long = (Me.Chk_long.Value = 1)
    ebstock_dess.opt_prof = (Me.Chk_prof.Value = 1)
    ebstock_dess.opt_rap = (Me.Chk_rap.Value = 1)
xlon = ebstock_dess.Longueur
xlar = ebstock_dess.Largeur
haut = ebstock_dess.Profondeur
rap = ebstock_dess.Rapport
If xlon + xlar + haut + rap > 0 Then
If ebstock_dess.opt_long And ebstock_dess.opt_larg And xlon > 0 And xlar > 0 Then
    If xlon > 0 And xlar > 0 Then
        surf = ebstock.volume / xlon
        haut = surf / xlar
        rap = xlar / haut
    Else
        haut = 0#
        rap = 0#
    End If
End If
If ebstock_dess.opt_long And ebstock_dess.opt_prof Then
    If xlon > 0 And haut > 0 Then
        surf = ebstock.volume / xlon
        xlar = surf / haut
        rap = xlar / haut
    Else
        xlar = 0#
        rap = 0#
    End If
End If
If ebstock_dess.opt_long And ebstock_dess.opt_rap Then
    If xlon > 0 And rap > 0 Then
        surf = ebstock.volume / xlon
        haut = Sqr(surf / rap)
        xlar = rap * haut
    Else
        haut = 0#
        xlar = 0#
    End If
End If
If ebstock_dess.opt_larg And ebstock_dess.opt_prof Then
    If xlar > 0 And haut > 0 Then
        surf = xlar * haut
        xlon = ebstock.volume / surf
        rap = xlar / haut
    Else
        xlon = 0#
        rap = 0#
    End If
End If
If ebstock_dess.opt_larg And ebstock_dess.opt_rap Then
    If xlar > 0 And rap > 0 Then
        haut = xlar / rap
        surf = xlar * haut
        xlon = ebstock.volume / surf
    Else
        haut = 0#
        xlon = 0#
    End If
End If
If ebstock_dess.opt_prof And ebstock_dess.opt_rap Then
    If haut > 0 And rap > 0 Then
        xlar = haut * rap
        surf = xlar * haut
        xlon = ebstock.volume / surf
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

ebstock_dess.Longueur = Round(xlon, 2)
ebstock_dess.Largeur = Round(xlar, 2)
ebstock_dess.Profondeur = Round(haut, 2)
ebstock_dess.Rapport = Round(rap, 2)
    Call dess_stock

End If
End Sub

Private Sub Tb_long_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_long.Text
    iSels = Tb_long.SelStart
    iSell = Tb_long.SelLength
    bKP = True
    chang_long = True
'    If Len(Tb_long.Text) <= Tb_long.MaxLength Then
'        KeyAscii = verif_car(Tb_long.Text, KeyAscii, "Saisie de la longueur", "R")
'    End If
End If
End Sub

Private Sub Tb_prof_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String

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
ebstock_dess.Profondeur = Round(haut, 2)
If Chk_prof.Value = 0 And Chk_prof.Enabled Then
    Chk_prof.Value = 1
End If
bkprof = True
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

Private Sub Tb_prof_Click()
Call sel_text(Tb_prof)

End Sub

Private Sub Tb_prof_GotFocus()
Call sel_text(Tb_prof)

End Sub

Private Sub Tb_prof_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_prof.Text
    iSels = Tb_prof.SelStart
    iSell = Tb_prof.SelLength
    bKP = True
    chang_prof = True
End Sub

Private Sub Tb_prof_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_prof.Text
    iSels = Tb_prof.SelStart
    iSell = Tb_prof.SelLength
    bKP = True
    chang_prof = True
'    If Len(Tb_prof.Text) <= Tb_prof.MaxLength Then
'        KeyAscii = verif_car(Tb_prof.Text, KeyAscii, "Saisie de la hauteur d'eau", "R")
'    End If
End If
End Sub

Private Sub Tb_Qav_Change()
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_Qav.Text, "Saisie du débit aval admissible", "I")
  If nom = "" Then
    Tb_Qav.Text = sval_champ
    Tb_Qav.SelStart = iSels
    Tb_Qav.SelLength = iSell
  End If
End If
'****
    ebstock.Qav = txtVersNum(Me.Tb_Qav.Text)
'    Call ini_ebstock_dess
'    ebstock.dessstock = ebstock_dess
'    Call init_form_dess
    Call reini_valeurs
     sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Qav_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Qav"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_Qav)
End Sub

Private Sub Tb_Qav_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Qav"
Call sel_text(Tb_Qav)
If change_coul Then
    Change_Couleur nom, 0
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_Qav_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_Qav.Text
    iSels = Tb_Qav.SelStart
    iSell = Tb_Qav.SelLength
    bKP = True
'   If Len(Tb_Qav.Text) <= Tb_Qav.MaxLength Then
'        KeyAscii = verif_car(Tb_Qav.Text, KeyAscii, "Saisie du débit aval admissible", "I")
'    End If
End If
End Sub
Private Function rec_vr(ByVal ipc As Double, ByVal ipr As Double) As Double
'1290 If ICAV < 0.5 Then Vr = 20 - (ICAV * 9)
'1300 If ICAV >= 0.5 And ICAV < 1 Then Vr = 15.5 - (ICAV - 0.5) * 1.25
'1310 If ICAV >= 1 And ICAV < 2 Then Vr = 13 - (ICAV - 1) * 4.3
'1320 If ICAV >= 2 And ICAV < 3 Then Vr = 8.7 - (ICAV - 2) * 2.2
'1330 If ICAV >= 3 And ICAV < 4 Then Vr = 6.5 - (ICAV - 3) * 1.5
'1340 If ICAV >= 4 And ICAV < 5 Then Vr = 5 - (ICAV - 4) * 1.1
'1350 If ICAV >= 5 And ICAV < 6 Then Vr = 3.9 - (ICAV - 5) * 0.8
'1360 If ICAV >= 6 And ICAV < 7 Then Vr = 3.1 - (ICAV - 6) * 0.6
'1370 If ICAV >= 7 And ICAV < 8 Then Vr = 2.5 - (ICAV - 7) * 0.5
'1380 If ICAV >= 8 And ICAV < 9 Then Vr = 2! - (ICAV - 8) * 0.4
'1390 If ICAV >= 9 And ICAV < 10 Then Vr = 1.6 - (ICAV - 9) * 0.2
'1400 If ICAV >= 10 Then Vr = 1.4 - (ICAV - 10) * 0.004
'1410 Vr = Vr * (icri / 15) ^ (1 / 2)
Dim list_TC(13, 2) As Double
Dim a As Double

Dim i As Integer

list_TC(1, 1) = 0
list_TC(2, 1) = 0.5
list_TC(3, 1) = 1
list_TC(4, 1) = 2
list_TC(5, 1) = 3
list_TC(6, 1) = 4
list_TC(7, 1) = 5
list_TC(8, 1) = 6
list_TC(9, 1) = 7
list_TC(10, 1) = 8
list_TC(11, 1) = 9
list_TC(12, 1) = 10
list_TC(13, 1) = 20
list_TC(1, 2) = 20
list_TC(2, 2) = 15.5
list_TC(3, 2) = 13
list_TC(4, 2) = 8.4
list_TC(5, 2) = 6.5
list_TC(6, 2) = 5
list_TC(7, 2) = 3.9
list_TC(8, 2) = 3.1
list_TC(9, 2) = 2.5
list_TC(10, 2) = 2#
list_TC(11, 2) = 1.6
list_TC(12, 2) = 1.4
list_TC(13, 2) = 1.36
'ipc = 11
i = 1
While ipc > list_TC(i, 1) And i < UBound(list_TC)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
a = (ipc - list_TC(i - 1, 1)) * (list_TC(i, 2) - list_TC(i - 1, 2)) / (list_TC(i, 1) - list_TC(i - 1, 1)) + list_TC(i - 1, 2)
a = a * (ipr / 15) ^ (1 / 2)
rec_vr = a



End Function
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, "" '"Dimensionnement d'un bassin de stockage"
End Sub

Private Sub Tb_Qav_LostFocus()
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_Qav", -1, txtVersNum(Tb_Qav.Text))
    If Not ok Then
        Tb_Qav.SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_rap_Change()
Dim xlon As Double, surf As Double, xlar As Double, haut As Double, rap As Double
Dim nom As String

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
ebstock_dess.Rapport = Round(rap, 2)
If Chk_rap.Value = 0 And Chk_rap.Enabled Then
    Chk_rap.Value = 1
End If
bkrap = True
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

Private Sub Tb_rap_Click()
Call sel_text(Tb_rap)

End Sub

Private Sub Tb_rap_GotFocus()
Call sel_text(Tb_rap)

End Sub

Private Sub Tb_rap_KeyDown(KeyCode As Integer, Shift As Integer)
    sval_champ = Tb_rap.Text
    iSels = Tb_rap.SelStart
    iSell = Tb_rap.SelLength
    bKP = True
    chang_rap = True
End Sub

Private Sub Tb_rap_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_rap.Text
    iSels = Tb_rap.SelStart
    iSell = Tb_rap.SelLength
    bKP = True
    chang_rap = True
'    If Len(Tb_rap.Text) <= Tb_rap.MaxLength Then
'        KeyAscii = verif_car(Tb_rap.Text, KeyAscii, "Saisie du rapport (largeur/hauteur d'eau)", "R")
'    End If
End If
End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub
