VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_lectfich 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche d'un ouvrage"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "Frm_lectfich.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   600
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_autre 
      Caption         =   "Autre étude"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   650
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1000
   End
   Begin VB.ComboBox Cb_fich 
      Height          =   315
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   650
      TabIndex        =   4
      Top             =   120
      Width           =   4600
   End
End
Attribute VB_Name = "Frm_lectfich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Private fich_texte As String
Public nomfich As String

Private Sub lect_fich()
    Select Case gnom_type
        Case Is = "versant"
            Call lect_bv
        Case Is = "chute"
            Call lect_chute
        Case Is = "pompe"
            Call lect_pompe
        Case Is = "conduite"
            Call lect_conduite
        Case Is = "decantation"
            Call lect_decant
        Case Is = "deversoir"
            Call lect_deversoir
        Case Is = "deversoiror"
            Call lect_deversoir_or
        Case Is = "retention"
            Call lect_ret
        Case Is = "siphon"
            Call lect_siphon
        Case Is = "stockage"
            Call lect_stock
    End Select
End Sub
Private Sub lect_stock()
Dim za As st_savstock
Dim za1 As st_savsto1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavstock
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_siphon()
Dim za As st_savsi
Dim za1 As st_savsi1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavsi
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_ret()
Dim za As st_savret
Dim za1 As st_savret1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavret
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_deversoir()
Dim za As st_savdo
Dim za1 As st_savdo1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdo
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_deversoir_or()
Dim za As st_savdoor
Dim za1 As st_savdoor1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdoor
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_decant()
Dim za As st_savdecant
Dim za1 As st_savdec1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdecant
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_conduite()
Dim za As st_savchute
Dim za1 As st_savch1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Function verif_etude(ByVal fsize As Double) As String
Dim za1 As st_hydrouti
Dim num As Integer
    lhFicDbf = FreeFile
    Open fich_lect_edit For Input Shared As #lhFicDbf ' Len = Len(za1)
    message = ""
    If fsize > 0 Then
    Input #lhFicDbf, num, za1.type
        If Trim(za1.type) <> "hydrouti" Then
            message = "Le fichier n'est pas de type HYDROUTI"
        End If
    End If
    Close #lhFicDbf
verif_etude = message
End Function
Private Sub lect_bv()
Dim za As st_save
Dim za1 As st_save1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        If Trim(za1.type) = gnom_type Then
            za = za1.stsave
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
'Cb_fich.Sorted
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_chute()
Dim za As st_savchute
Dim za1 As st_savch1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavch
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub
Private Sub lect_pompe()
Dim za As st_savpompe
Dim za1 As st_savpom1
    lhFicDbf = FreeFile
    Cb_fich.Clear
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavpo
        If Trim(za.type) = gnom_type Then
            Cb_fich.AddItem (za.nom)
        End If
   End If
Loop
Close #lhFicDbf
Cb_fich.Text = Cb_fich.list(0)
Cb_fich.Refresh
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If
End Sub



Private Sub Cb_fich_Change()
    Cb_fich.Text = fich_texte
End Sub

Private Sub Cb_fich_KeyDown(KeyCode As Integer, Shift As Integer)
    fich_texte = Cb_fich.Text
    Cb_fich.Text = fich_texte

End Sub

Private Sub Cb_ficht_KeyPress(KeyAscii As Integer)
    fich_texte = Cb_fich.Text
End Sub


Private Sub Cmd_autre_Click()
Dim reponse As Integer
Dim message As String, nomf As String, nomf1 As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim f As File
Dim d As Drive
Dim nom As String
cdlg1.DialogTitle = "Recherche d'un fichier "
cdlg1.FileName = ""
cdlg1.Filter = "Fichiers HYDROUTI (*.hyd)|*.hyd"
cdlg1.InitDir = ""
cdlg1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
cdlg1.ShowOpen
s = cdlg1.FileName
fsco = create_fs(s)
'    If fsco.dr_type = 1 Then
'        message = "Fichier sur disquette;" + Chr(13) + Chr(10) + "Vérifier que la disquette n'est pas protégée en écriture."
'        reponse = MsgBox(message, , "Saisie du nom de fichier")
'    End If
'    If fsco.dr_type = 4 Then
'        message = "Fichier sur CR-ROM;" + Chr(13) + Chr(10) + "Pas d'accés en écriture."
'        reponse = MsgBox(message, , "Saisie du nom de fichier")
'        nomf = ""
'    Else
    If fsco.lecteur <> "" And fsco.Chemin <> "" Then
        If Trim(fsco.nom) <> "" Then
'            If fsco.f_attr = 1 Or fsco.f_attr = 33 Then
'                message = "Fichier en lecture seule."
'                reponse = MsgBox(message, , "Saisie du nom de fichier")
'                nomf = ""
'            Else
                nomf = Trim(fsco.nomcomplet)
                nomf1 = Trim(fsco.nom)
'            End If
        Else
            nomf = ""
        End If
    Else
            nomf = ""
    End If
    If nomf <> "" Then
        If Right(nomf, 4) <> ".hyd" Then
            nomf = Left(nomf, Len(nomf) - 4) + ".hyd"
            nomf1 = Left(nomf1, Len(nomf1) - 4) + ".hyd"
        End If
        fich_lect_edit = nomf
        message = verif_etude(fsco.f_size)
        If message = "" Then
'            fich_lect = Left(fich_lect_edit, Len(fich_lect_edit) - 3) + "boh"
            fich_lect = chemin_app + Left(nomf1, Len(nomf1) - 3) + "boh"
            Call recup_fich(fich_lect, fich_lect_edit)
            Me.Caption = "Etude " + fich_lect_edit
            Call lect_fich
'            fich_lect = fich_lect_edit
        Else
            reponse = MsgBox(message, , "Saisie du nom de fichier")
        End If
    End If
End Sub

Private Sub Cmd_ok_Click()
    nomfich = Cb_fich.Text
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    Centre Me
    Set owner = MDIFrm_menu.rec_owner
    
If Not owner.fbassin Is Nothing Then
    gnom_type = owner.fbassin.nom_type
    For i = 0 To owner.fbassin.Cb_bassin.ListCount - 1
        Cb_fich.AddItem (owner.fbassin.Cb_bassin.list(i))
    Next
Else
'End If
If Not owner.fobjet Is Nothing Then
    gnom_type = owner.fobjet.nom_type
    Select Case UCase(owner.fobjet.Name)
        Case Is = "FRM_CHUTE"
            For i = 0 To owner.fobjet.Cb_chute.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_chute.list(i))
            Next
        Case Is = "FRM_POMPE"
            For i = 0 To owner.fobjet.Cb_pompe.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_pompe.list(i))
            Next
        Case Is = "FRM_CONDUITE"
            For i = 0 To owner.fobjet.Cb_conduite.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_conduite.list(i))
            Next
        Case Is = "FRM_SIPHON"
            For i = 0 To owner.fobjet.Cb_siphon.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_siphon.list(i))
            Next
        Case Is = "FRM_DO"
            For i = 0 To owner.fobjet.Cb_deversoir.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_deversoir.list(i))
            Next
        Case Is = "FRM_DO_OR"
            For i = 0 To owner.fobjet.Cb_deversoir.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_deversoir.list(i))
            Next
        Case Is = "FRM_DECANT"
            For i = 0 To owner.fobjet.Cb_decant.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_decant.list(i))
            Next
        Case Is = "FRM_STOCK"
            For i = 0 To owner.fobjet.Cb_stockage.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_stockage.list(i))
            Next
        Case Is = "FRM_RET"
            For i = 0 To owner.fobjet.Cb_retention.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_retention.list(i))
            Next
        Case Is = "FRM_BV2"
            For i = 0 To owner.fobjet.Cb_bassin.ListCount - 1
                Cb_fich.AddItem (owner.fobjet.Cb_bassin.list(i))
             Next
   End Select
End If
End If
    fich_texte = Cb_fich.list(0)
   Cb_fich.Text = Cb_fich.list(0)
End Sub

Private Sub Cmd_annul_Click()
    nomfich = ""
    fich_lect = ""
    Unload Me
End Sub
