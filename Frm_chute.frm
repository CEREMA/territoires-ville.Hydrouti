VERSION 5.00
Begin VB.Form Frm_chute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Chute"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_chute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.TextBox Tb_Qmax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4440
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1200
      Width           =   900
   End
   Begin VB.ComboBox Cb_chute 
      Height          =   315
      Left            =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   4000
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   4400
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Calcul de la chute"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.Frame Frm_Aval 
      Caption         =   "Conduite Aval "
      Height          =   2250
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Cmd_ava 
         Caption         =   "Courbe..."
         Height          =   255
         Left            =   2540
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Courbe de débit de la conduite aval"
         Top             =   1920
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
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   5
         Top             =   240
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
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   7
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Lb_uava 
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   32
         Top             =   1010
         Width           =   495
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Cote radier ZRav"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Lb_uava 
         Caption         =   "m"
         Height          =   255
         Index           =   3
         Left            =   2655
         TabIndex        =   24
         Top             =   1485
         Width           =   300
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Diamètre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Pente"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lb_uava 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   19
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Lb_uava 
         Caption         =   "1/10000"
         Height          =   255
         Index           =   1
         Left            =   2655
         TabIndex        =   18
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Lb_intava 
         Caption         =   "Coeff.  de Strickler"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   1455
      End
   End
   Begin VB.Frame Frm_Amont 
      Caption         =   "Conduite Amont "
      Height          =   2250
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Cmd_amo 
         Caption         =   "Courbe..."
         Height          =   255
         Left            =   2540
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Courbe de débit de la conduite amont"
         Top             =   1920
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
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
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
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   3
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Lb_uam 
         Height          =   255
         Index           =   2
         Left            =   2655
         TabIndex        =   31
         Top             =   1010
         Width           =   615
      End
      Begin VB.Label Lb_uam 
         Caption         =   "m"
         Height          =   255
         Index           =   3
         Left            =   2655
         TabIndex        =   23
         Top             =   1485
         Width           =   300
      End
      Begin VB.Label Lb_intam 
         Caption         =   "Cote radier ZRam"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Lb_intam 
         Caption         =   "Diamètre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Lb_intam 
         Caption         =   "Pente"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Lb_uam 
         Caption         =   "mm"
         Height          =   255
         Index           =   0
         Left            =   2655
         TabIndex        =   13
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Lb_uam 
         Caption         =   "1/10000"
         Height          =   255
         Index           =   1
         Left            =   2655
         TabIndex        =   12
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Lb_intam 
         Caption         =   "Coeff.  de Strickler"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1005
         Width           =   1335
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Lb_temp 
      Caption         =   "Lb_temp"
      Height          =   375
      Left            =   1200
      TabIndex        =   37
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Lb_Qmax 
      Caption         =   "Débit"
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   1240
      Width           =   375
   End
   Begin VB.Label Lb_uqmax 
      Caption         =   "m3/s"
      Height          =   255
      Left            =   5400
      TabIndex        =   34
      Top             =   1240
      Width           =   495
   End
   Begin VB.Label Lb_chute 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_chute"
      Height          =   1350
      Left            =   3480
      TabIndex        =   30
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Lb_ava 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_ava"
      Height          =   1350
      Left            =   6600
      TabIndex        =   29
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Lb_amo 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_amont"
      Height          =   1350
      Left            =   360
      TabIndex        =   28
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Chute"
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusaves 
         Caption         =   "En&registrer sous..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusuppr 
         Caption         =   "&Supprimer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuprint 
         Caption         =   "Im&primer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter module"
      End
   End
End
Attribute VB_Name = "Frm_chute"
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
Private ch_texte As String
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
Private Sub Change_Couleur(nom As String, Index As Integer)
'Dim coul As ColorConstants, coulp As ColorConstants
'Dim Index1 As Integer
'Dim nom1 As String
'coulp = vbBlack
'coul = Couleur_Change
'nom1 = nom
'Select Case nom
'    Case Is = "Tb_amo"
'         nom1 = "Lb_intam"
'    Case Is = "Tb_ava"
'         nom1 = "Lb_intava"
'    Case Is = "Tb_Qmax"
'         nom1 = "Lb_Qmax"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intam"
'         Lb_intam(index_prec).ForeColor = coulp
'    Case Is = "Lb_intava"
'         Lb_intava(index_prec).ForeColor = coulp
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coulp
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coulp
'    Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intam"
'         Lb_intam(Index).ForeColor = coul
'    Case Is = "Lb_intava"
'         Lb_intava(Index).ForeColor = coul
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coul
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coul
'   Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coul
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
    Case Is = "Lb_intam"
         Tb_amo(Index).SetFocus
    Case Is = "Lb_intava"
         Tb_ava(Index).SetFocus
    Case Is = "Lb_Qmax"
         Tb_Qmax.SetFocus
    Case Is = "Frm_Amont"
         Tb_amo(0).SetFocus
   Case Is = "Frm_Aval"
         Tb_ava(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
   Case Is = "Lb_intam", "Tb_amo"
    Select Case Index
        Case Is = 0
        mes = IDhlp_ChuteConduiteAmont '"Conduite Amont"
        Case Is = 1
        mes = IDhlp_ChuteConduiteAmont '"Conduite Amont"
        Case Is = 2
        mes = IDhlp_ChuteConduiteAmont '"Conduite Amont"
        Case Is = 3
        mes = IDhlp_ChuteConduiteAmont ' "Conduite Amont" '"Profil"
    End Select
    Case Is = "Lb_intava", "Tb_ava"
    Select Case Index
        Case Is = 0
        mes = IDhlp_ChuteConduiteAval ' "Conduite Aval"
        Case Is = 1
        mes = IDhlp_ChuteConduiteAval ' "Conduite Aval"
        Case Is = 2
        mes = IDhlp_ChuteConduiteAval ' "Conduite Aval"
        Case Is = 3
        mes = IDhlp_ChuteConduiteAval ' "Conduite Aval" ' "Profil"
    End Select
     Case Is = "Frm_Amont"
        mes = IDhlp_ChuteConduiteAmont '"Conduite Amont"
   Case Is = "Frm_Aval"
        mes = IDhlp_ChuteConduiteAval '"Conduite Aval"
    Case Is = "Lb_Qmax", "Tb_Qmax"
        mes = IDhlp_ChuteRegard '"Regard" '"Etude du profil"
    Case Is = "Cmd_calcul"
        mes = IDhlp_ChuteEtudeProfil ' "3. Etude du profil"
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
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Private Sub retaille()
    Me.Left = owner.fcom.Width + owner.fcom.Left
    Me.Top = 0
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc)  ' 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub

Private Sub Cb_chute_Change()
    Cb_chute.Text = ch_texte
End Sub

Private Sub Cb_chute_KeyDown(KeyCode As Integer, Shift As Integer)
    ch_texte = Cb_chute.Text
    Cb_chute.Text = ch_texte

End Sub

Private Sub Cb_chute_KeyPress(KeyAscii As Integer)
    ch_texte = Cb_chute.Text
End Sub


Private Sub Cmd_ava_Click()
Call dessin_courbe_ava
End Sub
Private Sub Cmd_calcul_Click()
If ebchute.Rdav > ebchute.Rdam Then
    Call calcul_amont_aval
    ouv_sauve = True
Else
    MsgBox "la cote aval doit être inférieure à la cote amont.", vbExclamation

End If
End Sub
Private Sub lect_fich()
Dim za As st_savchute
Dim za1 As st_savch1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_chute.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = nom_type Then
            Cb_chute.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
ch_texte = Cb_chute.list(0)
Cb_chute.Text = Cb_chute.list(0)
Cb_chute.Refresh

Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub

Private Sub Cmd_amo_Click()
Call dessin_courbe_amo
'Me.Cmd_calcul.Enabled = True
End Sub
Private Sub dessin_courbe_amo()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = ebchute.dam / 1000#
    canal.Longueur = 5
    canal.pente = ebchute.iRadam / 10000#
    canal.rugosite = ebchute.Kam
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = ebchute.Rdav
      .radamo = ebchute.Rdav + 0.3 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, val(Me.Tb_Qmax), "Courbe débit conduite amont")
End Sub
Private Sub dessin_courbe_ava()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = ebchute.dav / 1000#
    canal.Longueur = 5
    canal.pente = ebchute.iradav / 10000#
    canal.rugosite = ebchute.kav
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = ebchute.Rdam - 0.3
      .radamo = ebchute.Rdam 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, val(Me.Tb_Qmax), "Courbe débit conduite aval")
End Sub

Private Sub Cmd_calcul_GotFocus()
'Dim nom As String
'Dim mes As String
'    nom = "Cmd_calcul"
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
End Sub

Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'une chute"
    Change_Couleur "Me", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Frm_Amont_Click()
Dim mes As String
Dim nom As String
nom = "Frm_Amont"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'owner.affich_aide Me.Name, "Chute Conduite Amont"
End Sub

Private Sub Frm_Aval_Click()
Dim mes As String
Dim nom As String
nom = "Frm_Aval"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
'owner.affich_aide Me.Name, "Chute Conduite Aval"
End Sub






Private Sub Lb_intam_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intam"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
'owner.affich_aide Me.Name, "Chute Conduite Amont"
End Sub

Private Sub Lb_intava_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intava"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
'owner.affich_aide Me.Name, "Chute Conduite Aval"
End Sub

Private Sub Lb_Qmax_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Qmax"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
'owner.affich_aide Me.Name, "Chute Débit"

End Sub




Private Sub m_quitter_Click()
    Unload Me
    Unload owner
End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then  '(Not ouv_sauve And Not save_fich) Then
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
    reponse = MsgBox("La chute n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une chute")
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
    reponse = MsgBox("La chute n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une chute")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une chute "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_chute_click
            End If
        Case Is = 7
'            Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une chute "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_chute_click
            End If
    End Select
Else
'    Cb_chute.Visible = True
            frmf.Label1.Caption = "Recherche d'une chute "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_chute_click
            End If
End If
Set frmf = Nothing
End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer, nb As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "chute"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE CHUTE"
FrmPrint.sstitre1 = "Paramètres"
FrmPrint.ssTitre2 = ""
FrmPrint.ssTitre3 = "Résultats intermédiaires"
FrmPrint.ssTitre4 = ""
Frm_imp.Type1 = "chute"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE CHUTE"
Frm_imp.sstitre1 = "Paramètres"
Frm_imp.ssTitre2 = ""
Frm_imp.ssTitre3 = "Résultats intermédiaires"
Frm_imp.ssTitre4 = ""
nb = (Tb_amo.count - 1) + 1
ReDim list_don1(nb, 5)
    list_don1(0, 1) = ""
    list_don1(0, 2) = Frm_Amont.Caption
    list_don1(0, 3) = ""
    list_don1(0, 4) = Frm_Aval.Caption
    list_don1(0, 5) = ""
For i = 0 To Tb_amo.count - 1
    list_don1(i + 1, 1) = Lb_intam(i).Caption
    list_don1(i + 1, 2) = Tb_amo(i).Text
    list_don1(i + 1, 3) = Lb_uam(i).Caption
    list_don1(i + 1, 4) = Tb_ava(i).Text
    list_don1(i + 1, 5) = Lb_uava(i).Caption
Next
list_int2 = rec_list(Lb_amo.Caption)
list_int3 = rec_list(Lb_ava.Caption)
list_don1 = complet_listd_don(list_don1, list_int2, list_int3)
list_int1 = complet_listd_int1(list_int1, list_int2, list_int3)
ReDim list_don2(0, 3)
    list_don2(0, 1) = Lb_Qmax.Caption
    list_don2(0, 2) = Tb_Qmax.Text
    list_don2(0, 3) = Lb_uqmax.Caption
list_resu1 = rec_list(Lb_chute.Caption)
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
End Sub
Private Function complet_listd_don(ByVal liste1 As Variant, ByVal liste2 As Variant, ByVal liste3 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
i = -1
ReDim liste(UBound(liste1) + 2, 5)
For j = 0 To UBound(liste1)
    i = i + 1
    liste(i, 1) = liste1(j, 1)
    liste(i, 2) = liste1(j, 2)
    liste(i, 3) = liste1(j, 3)
    liste(i, 4) = liste1(j, 4)
    liste(i, 5) = liste1(j, 5)
Next
'i = i + 1
'liste(i, 1) = ""
'liste(i, 2) = ""
'liste(i, 3) = ""
i = i + 1
liste(i, 1) = liste2(0, 1)
liste(i, 2) = liste2(0, 2)
liste(i, 3) = liste2(0, 3)
liste(i, 4) = liste3(0, 2)
liste(i, 5) = liste3(0, 3)
i = i + 1
liste(i, 1) = liste2(1, 1)
liste(i, 2) = liste2(1, 2)
liste(i, 3) = liste2(1, 3)
liste(i, 4) = liste3(1, 2)
liste(i, 5) = liste3(1, 3)
complet_listd_don = liste
End Function
Private Function complet_listd_int1(ByVal liste1 As Variant, ByVal liste2 As Variant, ByVal liste3 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
        ReDim liste(2, 5)
        i = 0
        liste(i, 1) = ""
        liste(i, 2) = "Conduite amont"
        liste(i, 3) = ""
        liste(i, 4) = "Conduite aval"
        liste(i, 5) = ""
        i = i + 1
        liste(i, 1) = liste2(2, 1)
        liste(i, 2) = liste2(2, 2)
        liste(i, 3) = liste2(2, 3)
        liste(i, 4) = liste3(2, 2)
        liste(i, 5) = liste3(2, 3)
        i = i + 1
        liste(i, 1) = liste2(3, 1)
        liste(i, 2) = liste2(3, 2)
        liste(i, 3) = liste2(3, 3)
        liste(i, 4) = liste3(3, 2)
        liste(i, 5) = liste3(3, 3)
complet_listd_int1 = liste
End Function
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

Private Sub mnusaves_Click()
'    Me.Enabled = False
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'une chute "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde de la chute " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom de la chute (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
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
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'une chute")
        Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "chute"
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
        za.type = "chute"
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
        ch_texte = Trim(Tb_titre.Text)
        Cb_chute.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom de la chute n'est pas renseigné.", , "Sauvegarde d'une chute")
End If
 
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If

End Sub


Private Sub mnusuppr_Click()
Dim za As st_savchute
Dim za1 As st_savch1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_chute.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_chute.Text) + " va être supprimé .", 4, "Suppression d'une chute")
    If reponse = 6 Then  '6=oui,7=non
    save_fich = True
    ouv_sauve = False
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
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_chute.Text)) Then
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
    Me.Tb_amo(0).Text = "0"
    Me.Tb_amo(1).Text = "0"
    Me.Tb_amo(2).Text = "0"
    Me.Tb_ava(0).Text = "0"
    Me.Tb_ava(1).Text = "0"
    Me.Tb_ava(2).Text = "0"
    Me.Tb_amo(3).Text = "0.0"
    Me.Tb_ava(3).Text = "0.0"
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
Public Sub Cb_chute_click()
Dim za As st_savchute
Dim za1 As st_savch1
Call funlockb

'    Cb_chute.Visible = False
'    For i = 0 To Cb_chute.ListCount - 1
'        If Trim(Cb_chute.list(i)) = Trim(nom_ouvrage) Then
'            ch_texte = Cb_chute.list(i)
'            Cb_chute.Text = Cb_chute.list(i)
'        End If
'    Next
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    ch_texte = Trim(nom_ouvrage)
    Cb_chute.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_chute.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebchute = za.chute
            Call ini_form
            Call reini_valeurs
'           Me.Cmd_del.Visible = True
            If Cmd_calcul.Enabled Then
                Call Cmd_calcul_Click
            End If
            
            save_fich = True
            ouv_sauve = False
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("La chute n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une chute")
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
Private Sub Form_Load()
ichar = 0
  okg = True
  Me.KeyPreview = True
    Call ini_tooltip_chute(Me)
    ouv_sauve = False
    save_fich = False
'    save_fich = True
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'    Me.mnusave.Enabled = False
'    Me.mnusaves.Enabled = False
'    Me.Mnuprint.Enabled = False
'    Me.mnusuppr.Enabled = False
'''''    owner.affich_aide Me.Name, "Chute"
'    nom_fich = chemin_app + "ouvrages.bin"
'    nom_fich = chemin_app + "etude.boa"
    nom_type = "chute"
    fen_titre = Me.Caption
'   lecture fichier
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Cb_chute.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
    Call debut
End Sub
Private Sub debut0()
    Cb_chute.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
Dim itab As Integer
    bKP = False
    sval_champ = ""
Call init_l_tab
 Call donne_focus(Me)
    Me.Tb_amo(0).Text = "0"
    Me.Tb_amo(1).Text = "0"
    Me.Tb_amo(2).Text = "0"
    Me.Tb_ava(0).Text = "0"
    Me.Tb_ava(1).Text = "0"
    Me.Tb_ava(2).Text = "0"
    Me.Tb_amo(3).Text = "0.0"
    Me.Tb_ava(3).Text = "0.0"
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
    Frm_desprint.UC_graphique1.Height = 4500
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
'    owner.fdessin.UC_graphique1.reinit 7, "Arial"
'    owner.fdessin.UC_graphique1.init_title
'    owner.fdessin.UC_graphique1.init_titleh ""
'    owner.fdessin.UC_graphique1.init_titleb ""
    Call reini_valeurs
    Call ini_ebchute
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub ini_form()
    Me.Tb_amo(0).Text = rempl_virgule(Format(ebchute.dam, "###0"))
    Me.Tb_amo(1).Text = rempl_virgule(Format(ebchute.iRadam, "###0"))
    Me.Tb_amo(2).Text = rempl_virgule(Format(ebchute.Kam, "###0"))
    Me.Tb_ava(0).Text = rempl_virgule(Format(ebchute.dav, "###0"))
    Me.Tb_ava(1).Text = rempl_virgule(Format(ebchute.iradav, "###0"))
    Me.Tb_ava(2).Text = rempl_virgule(Format(ebchute.kav, "###0"))
    Me.Tb_amo(3).Text = rempl_virgule(Format(ebchute.Rdav, "##0.00"))
    Me.Tb_ava(3).Text = rempl_virgule(Format(ebchute.Rdam, "##0.00"))
    Me.Tb_Qmax.Text = rempl_virgule(Format(ebchute.Qmax, "#0.000"))
End Sub
Private Sub init_graph(ByRef uc_g As UC_graphique)
Dim ok As Boolean
Dim ecx As Double
Dim i As Integer
ok = False
uc_g.graphique_clear
uc_g.reinit 7, "Arial"
uc_g.init_titleh ""
uc_g.init_titleb ""
uc_g.init_arrondi_X 2
uc_g.init_arrondi_y 3
uc_g.init_MinX -2#
uc_g.init_MaxX ebchute.tron_amo.conduit.Longueur + ebchute.Long + ebchute.tron_ava.conduit.Longueur
uc_g.init_EchXn 1
ecx = uc_g.lire_EchXn()
uc_g.init_MaxY ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre + 0.5
uc_g.init_MinY Int(ebchute.tron_ava.radava) - 0.5
uc_g.init_EchYn 1
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    frm_menu.Enabled = True
    ouv_sauve = False
    Unload Frm_desprint
    Unload owner.fdessin
    owner.recharge_commentaire
End Sub

Private Sub MnuQuit_Click()
    Unload Me
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
Public Sub reini_valeurs()
' impression false
  Me.mnuprint.Enabled = False
owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Call ini_lbresu
     If ebchute.dam > 0 And ebchute.iRadam > 0 And ebchute.Kam > 0 _
        And ebchute.Rdav > 0 Then
        Me.Cmd_amo.Enabled = True
    Else
        Me.Cmd_amo.Enabled = False
    End If
     If ebchute.dav > 0 And ebchute.iradav > 0 And ebchute.kav > 0 _
        And ebchute.Rdam > 0 Then
        Me.Cmd_ava.Enabled = True
    Else
        Me.Cmd_ava.Enabled = False
    End If
   If ebchute.dam > 0 And ebchute.iRadam > 0 And ebchute.Kam > 0 _
        And ebchute.dav > 0 And ebchute.iradav > 0 And ebchute.kav > 0 _
        And ebchute.Rdav > 0 And ebchute.Rdam > 0 And ebchute.Qmax > 0 Then
        Me.Cmd_calcul.Enabled = True

    Else
        Me.Cmd_calcul.Enabled = False

    End If
    ouv_sauve = True

End Sub
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
    
    
    Select Case Index
        Case Is = 0
            ebchute.dam = txtVersNum(Me.Tb_amo(0).Text)
        Case Is = 1
            ebchute.iRadam = txtVersNum(Me.Tb_amo(1).Text)
        Case Is = 2
            ebchute.Kam = txtVersNum(Me.Tb_amo(2).Text)
        Case Is = 3
            ebchute.Rdav = txtVersNum(Me.Tb_amo(3).Text)
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
'owner.affich_aide Me.Name, "Chute Conduite Amont"

End Sub


Private Sub Tb_amo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_amo(Index).Text
    iSels = Tb_amo(Index).SelStart
    iSell = Tb_amo(Index).SelLength
'    If Len(Tb_amo(Index).Text) <= Tb_amo(Index).MaxLength Then
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
    
    Select Case Index
        Case Is = 0
            ebchute.dav = txtVersNum(Me.Tb_ava(0).Text)
        Case Is = 1
            ebchute.iradav = txtVersNum(Me.Tb_ava(1).Text)
        Case Is = 2
            ebchute.kav = txtVersNum(Me.Tb_ava(2).Text)
        Case Is = 3
            ebchute.Rdam = txtVersNum(Me.Tb_ava(3).Text)
    End Select
    Call reini_valeurs
    sval_champ = ""
    bKP = False
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
'owner.affich_aide Me.Name, "Chute Conduite Amont"
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
'owner.affich_aide Me.Name, "Chute Conduite Aval"
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
'owner.affich_aide Me.Name, "Chute Conduite Aval"

End Sub


Private Sub Tb_ava_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_ava(Index).Text
    iSels = Tb_ava(Index).SelStart
    iSell = Tb_ava(Index).SelLength
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

Private Sub Tb_Qmax_Change()
Dim nom As String

If bKP Then
             nom = verif_cart0(Tb_Qmax.Text, "Saisie débit ", "R")
  If nom = "" Then
    Tb_Qmax.Text = sval_champ
    Tb_Qmax.SelStart = iSels
    Tb_Qmax.SelLength = iSell
  End If
End If
'****
    ebchute.Qmax = txtVersNum(Me.Tb_Qmax.Text)
    Call reini_valeurs

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
DoEvents
Me.Show
Call sel_text(Tb_Qmax)
'owner.affich_aide Me.Name, "Chute Débit"
End Sub

Private Sub Tb_Qmax_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Qmax"
Call sel_text(Tb_Qmax)
'Tb_Qmax.SelStart = 0
'Tb_Qmax.SelLength = Len(Tb_Qmax.Text)
If change_coul Then
    Change_Couleur nom, 0
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End If
'owner.affich_aide Me.Name, "Chute Débit"

End Sub


Private Sub Tb_Qmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Qmax.Text
    iSels = Tb_Qmax.SelStart
    iSell = Tb_Qmax.SelLength
'    If Len(Tb_Qmax.Text) <= Tb_Qmax.MaxLength Then
'        KeyAscii = verif_car(Tb_Qmax.Text, KeyAscii, "Saisie débit ", "R")
'    End If
End If
End Sub
Private Sub ini_lbresu()
'    Me.Lb_amo.BackColor = &H8000000B
    Me.Lb_amo.BorderStyle = 1
    Me.Lb_amo.Caption = ""
'    Me.Lb_ava.BackColor = &H8000000B
    Me.Lb_ava.BorderStyle = 1
    Me.Lb_ava.Caption = ""
'    Me.Lb_chute.BackColor = &H8000000B
    Me.Lb_chute.BorderStyle = 1
    Me.Lb_chute.Caption = ""
End Sub
Private Sub modi_res_cana()
'    Me.Lb_amo.BackColor = &H80000009
    Me.Lb_amo.BorderStyle = 1
'    Me.Lb_ava.BackColor = &H80000009
    Me.Lb_ava.BorderStyle = 1
End Sub
Private Sub modi_res_chute()
'    Me.Lb_chute.BackColor = &H80000009
    Me.Lb_chute.BorderStyle = 1
End Sub
Private Sub calcul_amont_aval()
Dim z1 As Double, z2 As Double, h1 As Double, h2 As Double, h0 As Double
Dim v1 As Double, x0 As Double, X As Double, g As Double
Dim sresult As String, sresult1 As String, sresult2 As String
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
Dim res_amo As debit_conduit
Dim res_ava As debit_conduit
Dim cana_ava As conduite
Dim qv As deb_vit, qvps_amo As deb_vit, qvps_ava As deb_vit
g = 9.81
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
      .radava = ebchute.Rdav
      .radamo = ebchute.Rdav + cana_amo.Longueur * cana_amo.pente '0.3 '
    End With
    ebchute.tron_amo = troamo
    qvps_amo = debvit_ps(ebchute.tron_amo.conduit)
    res_amo = calc_debit_tr(ebchute.tron_amo, ebchute.Qmax)
    cana_ava.Diametre = ebchute.dav / 1000#
    cana_ava.Longueur = 5
    cana_ava.pente = ebchute.iradav / 10000#
    cana_ava.rugosite = ebchute.kav
    cana_ava.typ = 2
    With troava
      .Absava = 0#
      .Absava = .Absava + cana_ava.Longueur
      .conduit = cana_ava
      .radamo = ebchute.Rdam
      .radava = ebchute.Rdam - cana_ava.Longueur * cana_ava.pente ' 0.3 '
    End With
    ebchute.tron_ava = troava
    qvps_ava = debvit_ps(ebchute.tron_ava.conduit)
    res_ava = calc_debit_tr(ebchute.tron_ava, ebchute.Qmax)
    h1 = res_amo.hauteur
    h2 = res_ava.hauteur
    v1 = res_amo.vitesse
    z1 = ebchute.tron_amo.radava + h1
    z2 = ebchute.tron_ava.radamo + h2
'    Me.Lb_debitam.Caption = "Débit PS " + Trim(Str(Round(qvps_amo.debit, 3))) + " m3/s"
'    Me.Lb_debitav.Caption = "Débit PS " + Trim(Str(Round(qvps_ava.debit, 3))) + " m3/s"
'    Me.Lb_Vitam.Caption = "Vitesse PS " + Trim(Str(Round(qvps_amo.vitesse, 2))) + " m/s"
'    Me.Lb_Vitav.Caption = "Vitesse PS " + Trim(Str(Round(qvps_ava.vitesse, 2))) + " m/s"
    Call modi_res_cana
    sresult = "  Débit pleine section = " + ajout_zero(Trim(str(Round(qvps_amo.debit, 3)))) + " m3/s"
    sresult1 = "  Débit pleine section = " + ajout_zero(Trim(str(Round(qvps_ava.debit, 3)))) + " m3/s"
    sresult = sresult + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(str(Round(qvps_amo.vitesse, 2)))) + " m/s"
    sresult1 = sresult1 + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(str(Round(qvps_ava.vitesse, 2)))) + " m/s"

    Call init_graph(owner.fdessin.UC_graphique1)
    Call init_graph(Frm_desprint.UC_graphique1)
   
    If res_amo.charge Then
'        Me.Lb_Hautam.Caption = "Conduite en charge"
       sresult = sresult + Chr(13) + Chr(13) + "   Conduite en charge"
    Else
'        Me.Lb_Hautam.Caption = " Hauteur    " + Trim(Str(Round(res_amo.hauteur, 2))) + " m"
        sresult = sresult + Chr(13) + Chr(13) + "   Hauteur  = " + ajout_zero(Trim(str(Round(res_amo.hauteur, 2)))) + " m"
        sresult = sresult + Chr(13) + "   Vitesse = " + ajout_zero(Trim(str(Round(res_amo.vitesse, 2)))) + " m/s"
        If res_ava.charge Then
'            Me.Lb_Hautav.Caption = "Conduite en charge"
           sresult1 = sresult1 + Chr(13) + Chr(13) + "   Conduite en charge"
        Else
           Call modi_res_chute
'            Me.Lb_Hautav.Caption = "Hauteur    " + Trim(Str(Round(res_ava.hauteur, 2))) + " m"
           sresult1 = sresult1 + Chr(13) + Chr(13) + "   Hauteur = " + ajout_zero(Trim(str(Round(res_ava.hauteur, 2)))) + " m"
           sresult1 = sresult1 + Chr(13) + "   Vitesse = " + ajout_zero(Trim(str(Round(res_ava.vitesse, 2)))) + " m/s"
            h0 = z1 - z2
'            Me.Lb_Haut.Caption = "Dénivelée du liquide   " + Trim(Str(Round(h0, 2))) + " m"
'           sresult2 = "  Dénivelée = " + ajout_zero(Trim(Str(Round(h0, 2)))) + " m"
            ebchute.h0 = Round(h0, 2)
            Dim hc As Double, he As Double, le As Double
            he = 0: le = 0
            hc = long_chute(ebchute, res_amo, res_ava, he, le)
            h0 = ebchute.h0
'            sresult2 = "  Réduction de la hauteur d'eau = 0"
'            sresult2 = "  Hauteur initiale  = " + ajout_zero(Trim(Str(Round(he, 2)))) + " m "
'            sresult2 = sresult2 + Chr(13) + "   Longueur de variation = " + ajout_zero(Trim(Str(Round(le, 2)))) + " m "
'            sresult2 = sresult2 + Chr(13) + Chr(13) + "   Dénivelée = " + ajout_zero(Trim(Str(Round(h0, 2)))) + " m"
            ebchute.h0 = Round(h0, 2)
            x0 = ((h0 * v1 ^ 2) / g) ^ 0.5
            X = x0 * 2#
            X = hc
'            Me.Lb_Long.Caption = "Longueur du dispositif   " + Trim(Str(Round(X, 2))) + " m"
           sresult2 = "  Longueur du dispositif = " + ajout_zero(Trim(str(Round(X, 2)))) + " m"
            sresult2 = sresult2 + Chr(13) + "   Dénivelée = " + ajout_zero(Trim(str(Round(h0, 2)))) + " m"
            sresult2 = sresult2 + Chr(13) + Chr(13) + "   Hauteur initiale = " + ajout_zero(Trim(str(Round(he, 2)))) + " m "
            sresult2 = sresult2 + Chr(13) + "   Longueur de variation = " + ajout_zero(Trim(str(Round(le, 2)))) + " m "
            ebchute.Long = Round(X, 2)
            Me.Cmd_calcul.Enabled = False
            'impression true
            Me.mnuprint.Enabled = True
            Me.Lb_chute.Caption = sresult2
            Call dess_chute(owner.fdessin.UC_graphique1, h1, h2, v1)
            Call dess_chute1(owner.fdessin.UC_graphique1, ebchute, res_amo, res_ava)

            Call dess_chute(Frm_desprint.UC_graphique1, h1, h2, v1)
            Call dess_chute1(Frm_desprint.UC_graphique1, ebchute, res_amo, res_ava)
        End If
   End If
    Me.Lb_amo.Caption = sresult
    Me.Lb_ava.Caption = sresult1
End Sub
Public Sub dess_chute(ByRef uc_g As UC_graphique, ByVal h1 As Double, ByVal h2 As Double, ByVal v0 As Double)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim xy(11, 2) As Double, X As Double, dx As Double, Y As Double, x0 As Double, y0 As Double
x0 = ebchute.Long / 2
y0 = ebchute.h0 / 2

dx = x0 / 10
For i = 1 To 11
X = dx * (i - 1)
xy(i, 1) = X
Y = (X / v0) ^ 2 * 9.81 / 2
xy(i, 2) = Y
Next
'dessin des lignes epaisses grises
uc_g.redef_drwidth 10
xam = ebchute.tron_amo.Absamo
yam = ebchute.tron_amo.radamo
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_amo.radava
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, 0, -100
xam = ebchute.tron_amo.Absamo
yam = ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_amo.radava + ebchute.tron_amo.conduit.Diametre
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, 0, 100
xam = ebchute.tron_amo.Absava
yam = ebchute.tron_amo.radava
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_ava.radamo + (ebchute.Long * ebchute.tron_ava.conduit.pente)
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, 100, 0
xam = ebchute.tron_amo.Absava
yam = ebchute.tron_ava.radamo + (ebchute.Long * ebchute.tron_ava.conduit.pente)
xav = xam + ebchute.Long + ebchute.tron_ava.conduit.Longueur
yav = ebchute.tron_ava.radava
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, 0, -100
xam = ebchute.tron_amo.Absava + ebchute.Long
yam = ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre + 0.3
xav = ebchute.tron_amo.Absava + ebchute.Long
yav = ebchute.tron_ava.radamo + ebchute.tron_ava.conduit.Diametre
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, -100, 0
xam = ebchute.tron_amo.Absava + ebchute.Long
yam = ebchute.tron_ava.radamo + ebchute.tron_ava.conduit.Diametre
xav = xam + ebchute.tron_ava.conduit.Longueur
yav = ebchute.tron_ava.radava + ebchute.tron_ava.conduit.Diametre
uc_g.dess_ligndec xam, yam, xav, yav, couleur.gris_clair, 0, 100
uc_g.redef_drwidth 2
'dessin des contours
xam = ebchute.tron_amo.Absamo
yam = ebchute.tron_amo.radamo
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_amo.radava
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebchute.tron_amo.Absamo
yam = ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_amo.radava + ebchute.tron_amo.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebchute.tron_amo.Absava
yam = ebchute.tron_amo.radava
xav = ebchute.tron_amo.Absava
yav = ebchute.tron_ava.radamo + (ebchute.Long * ebchute.tron_ava.conduit.pente)
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebchute.tron_amo.Absava
yam = ebchute.tron_ava.radamo + (ebchute.Long * ebchute.tron_ava.conduit.pente)
xav = xam + ebchute.Long + ebchute.tron_ava.conduit.Longueur
yav = ebchute.tron_ava.radava
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebchute.tron_amo.Absava + ebchute.Long
yam = ebchute.tron_amo.radamo + ebchute.tron_amo.conduit.Diametre + 0.3
xav = ebchute.tron_amo.Absava + ebchute.Long
yav = ebchute.tron_ava.radamo + ebchute.tron_ava.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = ebchute.tron_amo.Absava + ebchute.Long
yam = ebchute.tron_ava.radamo + ebchute.tron_ava.conduit.Diametre
xav = xam + ebchute.tron_ava.conduit.Longueur
yav = ebchute.tron_ava.radava + ebchute.tron_ava.conduit.Diametre
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin des lignes d'eau
uc_g.redef_drwidth 3
'xam = ebchute.tron_amo.Absamo
'yam = ebchute.tron_amo.radamo + h1
'xav = ebchute.tron_amo.Absava
'yav = ebchute.tron_amo.radava + h1
'uc_g.dess_lign xam, yam, xav, yav, couleur.rouge, 2
'xam = ebchute.tron_amo.Absava + ebchute.Long
'yam = ebchute.tron_ava.radamo + h2
'xav = xam + ebchute.tron_ava.conduit.Longueur
'yav = ebchute.tron_ava.radava + h2
'uc_g.dess_lign xam, yam, xav, yav, couleur.rouge, 2

'xam = ebchute.tron_amo.Absava
'yam = ebchute.tron_amo.radava + h1
'x0 = xam
'y0 = yam
'For i = 1 To 11
'    X = xam + xy(i, 1)
'    Y = yam - xy(i, 2)
'    uc_g.dess_lign x0, y0, X, Y, couleur.rouge, 2
'    x0 = X
'    y0 = Y
'Next
'dy = Y - yam
'xam = X
'yam = Y
'For i = 1 To 11
'    X = xam + xy(i, 1)
'
'    Y = yam + dy + xy(12 - i, 2)
'    uc_g.dess_lign x0, y0, X, Y, couleur.rouge, 2
'    x0 = X
'    y0 = Y
'Next

'cotation
uc_g.redef_drwidth 1
uc_g.dess_coth_text ebchute.tron_amo.Absava, ebchute.tron_ava.radamo + (ebchute.Long * ebchute.tron_ava.conduit.pente), _
ebchute.tron_amo.Absava + ebchute.Long, ebchute.tron_ava.radamo + ebchute.tron_ava.conduit.Diametre, ajout_zero(Trim(str(ebchute.Long))) + " m", couleur_noir
'uc_g.dess_cotv_texte ebchute.tron_amo.Absava, ebchute.tron_amo.radava + h1, _
'ebchute.tron_amo.Absava + ebchute.Long, ebchute.tron_ava.radamo + h2, ajout_zero(Trim(Str(ebchute.h0))) + " m ", couleur_noir
uc_g.dess_cotv_texte ebchute.tron_amo.Absava, ebchute.tron_ava.radamo + h2 + ebchute.h0, _
ebchute.tron_amo.Absava + ebchute.Long, ebchute.tron_ava.radamo + h2, ajout_zero(Trim(str(ebchute.h0))) + " m ", couleur_noir
End Sub
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'une chute"
End Sub


Private Sub Tb_Qmax_LostFocus()
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_Qmax", -1, txtVersNum(Tb_Qmax.Text))
    If Not ok Then
        Tb_Qmax.SetFocus
        DoEvents
    End If
    okg = True
End If
End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub
Private Sub sel_text(tb_objet As TextBox)
    tb_objet.SelStart = 0
    
    tb_objet.SelLength = Len(tb_objet.Text)


End Sub
