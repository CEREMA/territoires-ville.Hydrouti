VERSION 5.00
Begin VB.Form Frm_conduite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Conduite"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_conduite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9825
   Begin hydrouti.UC_graphique UC_graphiquec 
      Height          =   2535
      Left            =   4200
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4471
   End
   Begin VB.CommandButton Cmd_cond 
      Caption         =   "Graphique"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Courbe de débit de la conduite "
      Top             =   2400
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Vérifications à partir du débit"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Tb_Qmax 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2760
      Width           =   900
   End
   Begin VB.ComboBox Cb_conduite 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4000
   End
   Begin VB.Frame Frm_conduite 
      Caption         =   "Conduite"
      Height          =   2055
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   3735
      Begin VB.TextBox Tb_cond 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1305
         Width           =   960
      End
      Begin VB.TextBox Tb_cond 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   2
         Top             =   855
         Width           =   960
      End
      Begin VB.TextBox Tb_cond 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   1
         Top             =   495
         Width           =   960
      End
      Begin VB.Label Lb_ucond 
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   11
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Lb_ucond 
         Caption         =   "1/10000"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Lb_ucond 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   9
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Lb_cond 
         Caption         =   "Coeff. de  Strickler"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Lb_cond 
         Caption         =   "Pente"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Lb_cond 
         Caption         =   "Diamètre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1395
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Lb_resu 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_resu"
      Height          =   855
      Left            =   5760
      TabIndex        =   18
      Top             =   3200
      Width           =   3735
   End
   Begin VB.Label Lb_conduite 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_conduite"
      Height          =   855
      Left            =   360
      TabIndex        =   16
      Top             =   3200
      Width           =   3735
   End
   Begin VB.Label Lb_uqmax 
      Caption         =   "m3/s"
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   2805
      Width           =   495
   End
   Begin VB.Label Lb_Qmax 
      Caption         =   "Débit"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2805
      Width           =   735
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Conduite"
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
      Begin VB.Menu mnuprint 
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
Attribute VB_Name = "Frm_conduite"
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
Private list_don1() As Variant
Private list_int1() As Variant
Private list_don2() As Variant
Private list_int2() As Variant
Private list_don3() As Variant
Private list_int3() As Variant
Private list_resu1() As Variant
Private co_texte As String
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
'    Case Is = "Tb_cond"
'         nom1 = "Lb_cond"
'    Case Is = "Tb_Qmax"
'         nom1 = "Lb_Qmax"
'End Select
'Select Case label_prec
'    Case Is = "Lb_cond"
'         Lb_cond(index_prec).ForeColor = coulp
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coulp
'    Case Is = "Frm_conduite"
'         Frm_conduite.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_cond"
'         Lb_cond(Index).ForeColor = coul
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coul
'    Case Is = "Frm_conduite"
'         Frm_conduite.ForeColor = coul
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
    Case Is = "Lb_cond"
         Tb_cond(Index).SetFocus
    Case Is = "Lb_Qmax"
         Tb_Qmax.SetFocus
    Case Is = "Frm_conduite"
         Tb_cond(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_cond", "Tb_cond", "Frm_conduite"
        mes = IDhlp_ConduiteDimensionnement  '"Dimensionnement d'une conduite"
    Case Is = "Lb_Qmax", "Tb_Qmax"
        mes = IDhlp_ConduiteDimensionnement  '"Dimensionnement d'une conduite"
End Select
mes_prec = mes
Rec_Mes = mes
End Function

Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Public Sub ini_ebchute()
    ebchute.dam = 0
    ebchute.iRadam = 0
    ebchute.Kam = 0
    ebchute.dav = 0
    ebchute.kav = 0
    ebchute.iradav = 0
    ebchute.Rdav = 0#
    ebchute.Rdam = 0#
    ebchute.Qmax = 0#
    ebchute.h0 = 5#
    ebchute.Long = 5#
End Sub
Private Sub ini_form()
    Me.Tb_cond(0).Text = rempl_virgule(Format(ebchute.dam, "###0"))
    Me.Tb_cond(1).Text = rempl_virgule(Format(ebchute.iRadam, "###0"))
    Me.Tb_cond(2).Text = rempl_virgule(Format(ebchute.Kam, "###0"))
    Me.Tb_Qmax.Text = rempl_virgule(Format(ebchute.Qmax, "#0.000"))
End Sub
Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant
l0 = Array(0)
l1 = Array(0, "TB_cond", "TB_Qmax")
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1))
list_tb = Array(l0, l1)

End Sub
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, ""
End Sub
Public Sub retailler()
retaille

End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_savchute
Dim za1 As st_savch1
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
            za = za1.stsavch
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'une conduite")
        Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "conduite"
            za.nom = Tb_titre.Text
            za.chute = ebchute
            za1.stsavch = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        za.type = "conduite"
        za.nom = Tb_titre.Text
        za.chute = ebchute
        za1.stsavch = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        co_texte = Trim(Tb_titre.Text)
        Cb_conduite.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom de la conduite n'est pas renseigné.", , "Sauvegarde d'une chute")
End If
 

Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Private Sub ini_lbresu()
'    Me.Lb_conduite.BackColor = &H8000000B
    Me.Lb_conduite.BorderStyle = 1
    Me.Lb_conduite.Caption = ""
'    Me.Lb_resu.BackColor = &H8000000B
    Me.Lb_resu.BorderStyle = 1
    Me.Lb_resu.Caption = ""
End Sub
Private Sub ini_lb_lbresu()
'    Me.Lb_resu.BackColor = &H8000000B
    Me.Lb_resu.BorderStyle = 1
    Me.Lb_resu.Caption = ""
End Sub
Private Sub lect_fich()
Dim za As st_savchute
Dim za1 As st_savch1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_conduite.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = nom_type Then
            Cb_conduite.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
co_texte = Cb_conduite.list(0)
Cb_conduite.Text = Cb_conduite.list(0)
Cb_conduite.Refresh
 
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
Public Sub Mquit()
    m_quitter_Click
End Sub
Private Sub modi_res_resu()
'    Me.Lb_resu.BackColor = &H80000009
    Me.Lb_resu.BorderStyle = 1
End Sub
Private Sub modi_res_conduite()
'    Me.Lb_conduite.BackColor = &H80000009
    Me.Lb_conduite.BorderStyle = 1
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Public Sub reini_valeurs()
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Me.UC_graphiquec.graphique_clear
    Call ini_lbresu
   ' impression false
    Me.mnuprint.Enabled = False
   If ebchute.dam > 0 And ebchute.iRadam > 0 And ebchute.Kam > 0 Then
           Call Cmd_cond_Click
        Me.Cmd_cond.Enabled = True
        If ebchute.Qmax > 0 Then
            Me.Cmd_calcul.Enabled = True
            Call Cmd_calcul_Click
        Else
            Me.Cmd_calcul.Enabled = False
        End If
    Else
        Me.Cmd_calcul.Enabled = False
        Me.Cmd_cond.Enabled = False
    End If
    ouv_sauve = True
End Sub
Private Sub retaille()
    Me.Left = owner.fcom.Width + owner.fcom.Left
    Me.Top = 0
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc) ' 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub


Private Sub Cmd_calcul_Click()
Call calcul_amont_aval
ouv_sauve = True
End Sub

Private Sub Cmd_cond_Click()
Call dessin_courbe_débit
Me.mnuprint.Enabled = True
ouv_sauve = True
End Sub


Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""
    Change_Couleur "Me", 0

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Form_Load()
     okg = True
      Me.KeyPreview = True
    Call ini_tooltip_conduite(Me)
    ouv_sauve = False
    save_fich = False
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'    Me.mnusave.Enabled = False
'    Me.mnusaves.Enabled = False
'    Me.Mnuprint.Enabled = False
'    Me.mnusuppr.Enabled = False
'''    owner.affich_aide Me.Name, "Conduite"
'    nom_fich = chemin_app + "conduites.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "conduite"
    fen_titre = Me.Caption
'   lecture fichier
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Cb_conduite.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
    Call debut0
'    Call debut
End Sub
Private Sub debut0()
    Cb_conduite.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub calcul_amont_aval()
'Dim g As Double
Dim sresult As String
Dim troamo As troncon
Dim cana_amo As conduite
Dim res_amo As debit_conduit
Dim qvps_amo As deb_vit
'g = 9.81
' conduite amont -> troncon amont
    cana_amo.Diametre = ebchute.dam / 1000#
    cana_amo.Longueur = 5
    cana_amo.pente = ebchute.iRadam / 10000#
    cana_amo.rugosite = ebchute.Kam
    cana_amo.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + cana_amo.Longueur
      .conduit = cana_amo
      .radava = 100#
      .radamo = 100.3 'cana_amo.Longueur * cana_amo.pente
    End With
    ebchute.tron_amo = troamo
    qvps_amo = debvit_ps(ebchute.tron_amo.conduit)
    res_amo = calc_debit_tr(ebchute.tron_amo, ebchute.Qmax)
'    h1 = res_amo.hauteur
'    v1 = res_amo.vitesse
'    z1 = ebchute.tron_amo.radava + h1
'    z2 = ebchute.tron_ava.radamo + h2
    Call modi_res_resu
'    sresult = "  Débit pleine section = " + ajout_zero(Trim(Str(Round(qvps_amo.debit, 3)))) + " m3/s"
'    sresult = sresult + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(Str(Round(qvps_amo.vitesse, 2)))) + " m/s"
    
   
    If res_amo.charge Then
       sresult = "   Conduite en charge"
'       sresult = sresult + Chr(13) + Chr(13) + "   Conduite en charge"
        ' impression false
'                    Me.mnuprint.Enabled = False
    Else
'        sresult = sresult + Chr(13) + Chr(13) + "   Hauteur  = " + ajout_zero(Trim(Str(Round(res_amo.hauteur, 2)))) + " m"
        sresult = "   Hauteur  = " + ajout_zero(Trim(str(Round(res_amo.hauteur, 2)))) + " m"
        sresult = sresult + Chr(13) + "   Vitesse = " + ajout_zero(Trim(str(Round(res_amo.vitesse, 2)))) + " m/s"
        ' impression true
'                    Me.mnuprint.Enabled = True
   End If
    Me.Lb_resu.Caption = sresult
End Sub
Private Sub dessin_courbe_débit()
Dim sresult As String
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
'Dim res_amo As debit_conduit
Dim qv As deb_vit, qvps_amo As deb_vit, qvps_ava As deb_vit
   cana_amo.Diametre = ebchute.dam / 1000#
    cana_amo.Longueur = 50
    cana_amo.pente = ebchute.iRadam / 10000#
    cana_amo.rugosite = ebchute.Kam
    cana_amo.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + cana_amo.Longueur
      .conduit = cana_amo
      .radava = 100#
      .radamo = 100# + cana_amo.Longueur * cana_amo.pente
    End With
    ebchute.tron_amo = troamo
    qvps_amo = debvit_ps(ebchute.tron_amo.conduit)
'    res_amo = calc_debit_tr(ebchute.tron_amo, ebchute.Qmax)
    Call modi_res_conduite
    sresult = "  Débit pleine section = " + ajout_zero(Trim(str(Round(qvps_amo.debit, 3)))) + " m3/s"
    sresult = sresult + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(str(Round(qvps_amo.vitesse, 2)))) + " m/s"
    Call init_graph(Me.UC_graphiquec)
    Call dess_conduite(Me.UC_graphiquec, ebchute.tron_amo)
    owner.fdessin.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
    Call calc_courbe_debit_tr(owner.fdessin.UC_graphique1, ebchute.tron_amo)
    owner.fdessin.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
    owner.fdessin.UC_graphique1.dess_lign 0, qvps_amo.debit * 1000, ebchute.dam, qvps_amo.debit * 1000, couleur.orange, 1
    owner.fdessin.UC_graphique1.init_lbvh "l/s"
    owner.fdessin.UC_graphique1.init_lbhd "mn"
    Frm_desprint.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
    Call calc_courbe_debit_tr(Frm_desprint.UC_graphique1, ebchute.tron_amo)
    Frm_desprint.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
    Frm_desprint.UC_graphique1.dess_lign 0, qvps_amo.debit * 1000, ebchute.dam, qvps_amo.debit * 1000, couleur.orange, 1
    Frm_desprint.UC_graphique1.init_lbvh "l/s"
    Frm_desprint.UC_graphique1.init_lbhd "mn"
   Me.Lb_conduite.Caption = sresult
End Sub
Private Sub debut()
Dim itab As Integer
    bKP = False
    sval_champ = ""
 Call init_l_tab
 Call donne_focus(Me)
    Me.Tb_cond(0).Text = "0"
    Me.Tb_cond(1).Text = "0"
    Me.Tb_cond(2).Text = "0"
    Me.Tb_Qmax.Text = "0.0"
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Me.UC_graphiquec.graphique_clear
    Call reini_valeurs
    Call ini_ebchute
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub init_graph(ByRef uc_g As UC_graphique)
Dim ok As Boolean
Dim ecx As Double
Dim i As Integer
ok = False
uc_g.graphique_clear
uc_g.reinit 7, "Arial"
uc_g.redef_cadrs 0, 0, 0
uc_g.init_titleh ""
uc_g.init_titleb ""
'uc_g.init_arrondi_X 2
'uc_g.init_arrondi_y 3
uc_g.init_MinX -2#
uc_g.init_MaxX ebchute.tron_amo.conduit.Longueur + 2
uc_g.init_EchXn 1#
'ecx = uc_g.lire_EchXn()
uc_g.init_MaxY ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre + 0.5
uc_g.init_MinY Int(ebchute.tron_amo.radava) - 0.5
uc_g.init_EchYn 1#
  
End Sub
Private Sub dess_conduite(ByRef uc_g As UC_graphique, ByRef tr1 As troncon)
Dim tr As troncon
tr.Absamo = 0
tr.radamo = tr1.radamo '- tr1.radava
tr.Absava = tr1.Absava + tr.Absamo
tr.radava = tr1.radava '- tr1.radava
tr.conduit.Diametre = tr1.conduit.Diametre
Call dess_troncon_c(uc_g, tr, couleur.bleu)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("La conduite n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une conduite")
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

Private Sub Frm_conduite_Click()
Dim mes As String
Dim nom As String
nom = "Frm_conduite"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'owner.affich_aide Me.Name, "Chute Conduite Amont"
End Sub
Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub Lb_cond_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_cond"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_Qmax_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Qmax"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then 'Or (Not ouv_sauve And Not save_fich) Then
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
    reponse = MsgBox("La conduite n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une conduite")
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
    reponse = MsgBox("La conduite n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une conduite")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une conduite "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_conduite_click
            End If
        Case Is = 7
'            Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une conduite "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_conduite_click
            End If
    End Select
Else
'    Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une conduite "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_conduite_click
            End If
End If
Set frmf = Nothing
End Sub

Public Function lect_list(ByVal nom As String) As Variant
Select Case nom
Case Is = "list_don1"
    lect_list = list_don1
Case Is = "list_int1"
    lect_list = list_int1
Case Is = "list_resu1"
    lect_list = list_resu1
Case Is = "list_don2"
    lect_list = list_don2
Case Is = "list_int2"
    lect_list = list_int2
Case Is = "list_resu2"
    lect_list = list_resu2
Case Is = "list_don3"
    lect_list = list_don3
End Select
End Function
Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer, nb As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "conduite"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE CONDUITE"
FrmPrint.sstitre1 = "Paramètres"
FrmPrint.ssTitre2 = ""
FrmPrint.ssTitre3 = ""
Frm_imp.Type1 = "conduite"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE CONDUITE"
Frm_imp.sstitre1 = "Paramètres"
Frm_imp.ssTitre2 = ""
Frm_imp.ssTitre3 = ""
nb = (Tb_cond.count - 1) + 1
ReDim list_don1(nb, 3)
    list_don1(0, 1) = ""
    list_don1(0, 2) = Frm_conduite.Caption
    list_don1(0, 3) = ""
For i = 0 To Tb_cond.count - 1
    list_don1(i + 1, 1) = Lb_cond(i).Caption
    list_don1(i + 1, 2) = Tb_cond(i).Text
    list_don1(i + 1, 3) = Lb_ucond(i).Caption
Next
list_int2 = rec_list(Lb_conduite.Caption)
list_don1 = complet_listd_don(list_don1, list_int2)
If Trim(Lb_resu.Caption) <> "" And Trim(Lb_resu.Caption) <> "Conduite en charge" Then
    list_int3 = rec_list(Lb_resu.Caption)
    list_int1 = complet_listd_int1(list_int1, list_int3)
    FrmPrint.ssTitre3 = "Résultats intermédiaires"
    Frm_imp.ssTitre3 = "Résultats intermédiaires"
End If
ReDim list_don2(0, 3)
    list_don2(0, 1) = Lb_Qmax.Caption
    list_don2(0, 2) = Tb_Qmax.Text
    list_don2(0, 3) = Lb_uqmax.Caption
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
End Sub
Private Function complet_listd_don(ByVal liste1 As Variant, ByVal liste2 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
i = -1
ReDim liste(UBound(liste1) + 2, 3)
For j = 0 To UBound(liste1)
    i = i + 1
    liste(i, 1) = liste1(j, 1)
    liste(i, 2) = liste1(j, 2)
    liste(i, 3) = liste1(j, 3)
Next
'i = i + 1
'liste(i, 1) = ""
'liste(i, 2) = ""
'liste(i, 3) = ""
i = i + 1
liste(i, 1) = liste2(0, 1)
liste(i, 2) = liste2(0, 2)
liste(i, 3) = liste2(0, 3)
i = i + 1
liste(i, 1) = liste2(1, 1)
liste(i, 2) = liste2(1, 2)
liste(i, 3) = liste2(1, 3)
complet_listd_don = liste
End Function
Private Function complet_listd_int1(ByVal liste1 As Variant, ByVal liste2 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
        ReDim liste(2, 3)
        i = 0
        liste(i, 1) = ""
        liste(i, 2) = "Conduite "
        liste(i, 3) = ""
        i = i + 1
        liste(i, 1) = liste2(0, 1)
        liste(i, 2) = liste2(0, 2)
        liste(i, 3) = liste2(0, 3)
        i = i + 1
        liste(i, 1) = liste2(1, 1)
        liste(i, 2) = liste2(1, 2)
        liste(i, 3) = liste2(1, 3)
complet_listd_int1 = liste
End Function

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Public Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If

End Sub
Private Sub mnusaves_Click()
'    Me.Enabled = False
 'modif FO   ' If ProtectCheck(2) <> 0 Then End
   If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'une conduite "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde de la conduite " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom de la conduite (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub
Private Sub mnusuppr_Click()
Dim za As st_savchute
Dim za1 As st_savch1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_conduite.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_conduite.Text) + " va être supprimé .", 4, "Suppression d'une conduite")
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
            za = za1.stsavch
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_conduite.Text)) Then
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
    Call ini_ebchute
    Me.Tb_cond(0).Text = "0"
    Me.Tb_cond(1).Text = "0"
    Me.Tb_cond(2).Text = "0"
    Me.Tb_Qmax.Text = "0.0"
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
Private Sub Tb_cond_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_cond(Index).Text, "Saisie diamètre conduite ", "I")
            Case Is = 1
                nom = verif_cart0(Tb_cond(Index).Text, "Saisie pente conduite ", "I")
            Case Is = 2
                nom = verif_cart0(Tb_cond(Index).Text, "Saisie coefficient conduite ", "I")
        End Select
  If nom = "" Then
    Tb_cond(Index).Text = sval_champ
    Tb_cond(Index).SelStart = iSels
    Tb_cond(Index).SelLength = iSell
  End If
End If
'****

    Select Case Index
        Case Is = 0
            ebchute.dam = txtVersNum(Me.Tb_cond(0).Text)
        Case Is = 1
            ebchute.iRadam = txtVersNum(Me.Tb_cond(1).Text)
        Case Is = 2
            ebchute.Kam = txtVersNum(Me.Tb_cond(2).Text)
    End Select
    Call reini_valeurs
     sval_champ = ""
    bKP = False
End Sub

Private Sub Tb_cond_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_cond"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_cond(Index))

End Sub

Private Sub Tb_cond_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_cond"
Call sel_text(Tb_cond(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub
Private Sub Tb_cond_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then  'Or KeyAscii = 9
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_cond(Index).Text
    iSels = Tb_cond(Index).SelStart
    iSell = Tb_cond(Index).SelLength
'    If Len(Tb_cond(Index).Text) <= Tb_cond(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_cond(Index).Text, KeyAscii, "Saisie diamètre conduite ", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_cond(Index).Text, KeyAscii, "Saisie pente conduite ", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_cond(Index).Text, KeyAscii, "Saisie coefficient conduite ", "I")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_cond_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_cond", Index, txtVersNum(Tb_cond(Index).Text))
    If Not ok Then
        Tb_cond(Index).SetFocus
        DoEvents
    End If
    okg = True
End If
End Sub
Private Sub Tb_Qmax_Change()
Dim nom As String

'If sval_champ <> "" Then
If bKP Then
        nom = verif_cart0(Tb_Qmax.Text, "Saisie débit ", "R")
  If nom = "" Then
    Tb_Qmax.Text = sval_champ
    Tb_Qmax.SelStart = iSels
    Tb_Qmax.SelLength = iSell
    Tb_Qmax.SetFocus
'  End If
'End If
    Else
'****

    ebchute.Qmax = txtVersNum(Me.Tb_Qmax.Text)
    Call ini_lb_lbresu
   ' impression false
'    Me.mnuprint.Enabled = False
    If ebchute.Qmax > 0 Then
       If Me.Cmd_cond.Enabled Then
            Me.Cmd_calcul.Enabled = True
            Call Cmd_calcul_Click
            Call dessin_courbe_débit
        End If
    Else
        Me.Cmd_calcul.Enabled = False
       If Me.Cmd_cond.Enabled Then
            Call dessin_courbe_débit
        End If
    End If
  End If
End If
 sval_champ = ""
 bKP = False
End Sub

Private Sub Tb_Qmax_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Qmax"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call sel_text(Tb_Qmax)
End Sub

Private Sub Tb_Qmax_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Qmax"
Call sel_text(Tb_Qmax)
If change_coul Then
    Change_Couleur nom, 0
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_Qmax_KeyPress(KeyAscii As Integer)
Dim reponse As Integer
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_Qmax.Text
    iSels = Tb_Qmax.SelStart
    iSell = Tb_Qmax.SelLength
    bKP = True
'   If Len(Tb_Qmax.Text) <= Tb_Qmax.MaxLength Then
'        KeyAscii = verif_car(Tb_Qmax.Text, KeyAscii, "Saisie débit ", "R")
'    End If
End If
End Sub
Private Sub Cb_conduite_Change()
    Cb_conduite.Text = co_texte
End Sub
Public Sub Cb_conduite_click()
Dim za As st_savchute
Dim za1 As st_savch1
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    co_texte = Trim(nom_ouvrage)
    Cb_conduite.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_conduite.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebchute = za.chute
            Call ini_form
            Call reini_valeurs
'           Me.Cmd_del.Visible = True
            If ebchute.dam > 0 And ebchute.iRadam > 0 And ebchute.Kam > 0 Then
                Call Cmd_cond_Click
                If ebchute.Qmax > 0 Then
                    Call Cmd_calcul_Click
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
Private Sub Cb_conduite_KeyDown(KeyCode As Integer, Shift As Integer)
    co_texte = Cb_conduite.Text
    Cb_conduite.Text = co_texte

End Sub

Private Sub Cb_conduite_KeyPress(KeyAscii As Integer)
    co_texte = Cb_conduite.Text
End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub

