VERSION 5.00
Begin VB.Form Frm_recopie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changement de version BOHHA"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Frm_recopie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_quit 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   3480
      Width           =   1000
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "Frm_recopie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FileLength As Integer
Private lhFicDbf As Integer
Private lhFicDbf1 As Integer
Private Sub Cmd_Quit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Centre Me
     Label1.Caption = Chr(10) + "  La nouvelle version du programme n'utilise plus qu'un seul fichier " _
        + "qui regroupe l'ensemble des ouvrages ." + Chr(10) _
        + "  En quittant cette fenêtre les fichiers qui ont été créés sur le répertoire de travail dans une version précédente " _
        + "vont être recopiés dans un fichier ETUDE.BOA sur ce répertoire de travail et détruits." + Chr(10) _
        + "  Ce fichier pourra ensuite être renommé avec une extension .BOA " _
        + "et recopié sur un répertoire quelconque." + Chr(10) + Chr(10) _
        + "  Cette nouvelle version permet de travailler avec plusieurs études " _
        + "sauvegardées sur des répertoires différents." + Chr(10) + Chr(10)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call recopie_fich0
        MDIFrm_menu.m_etu.Enabled = True
End Sub
Public Sub recopie_bv0(ByVal nom1 As String)
Dim za As st_save
Dim za1 As st_save1
Dim nom As String
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.type = "versant"
        za1.stsave = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Private Sub init_etude0()
Dim nom As String
Dim za1 As st_hydrouti
nom = chemin_app + "etude.boh"
    lhFicDbf = FreeFile
    Open nom For Random Access Write As #lhFicDbf Len = Len(za1)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    za1.type = "hydrouti"
    FileLength = LOF(lhFicDbf) / Len(za1) + 1
    Put #lhFicDbf, FileLength, za1

    Close #lhFicDbf

End Sub
Public Sub recopie_chute0(ByVal nom1 As String)
Dim nom As String
Dim za As st_savchute
Dim za1 As st_savch1
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavch = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_siphon0(ByVal nom1 As String)
Dim nom As String
Dim za As st_savsi
Dim za1 As st_savsi1
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavsi = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_conduite0(ByVal nom1 As String)
Dim za As st_savchute
Dim za1 As st_savch1
Dim nom As String
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavch = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_decant0(ByVal nom1 As String)
Dim za As st_savdecant
Dim za1 As st_savdec1
Dim nom As String
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavdecant = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_deversoir0(ByVal nom1 As String)
Dim za As st_savdo
Dim za1 As st_savdo1
Dim nom As String
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
    za.type = "deversoir"
   If Not EOF(lhFicDbf) Then
        za1.stsavdo = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_ret0(ByVal nom1 As String)
Dim za As st_savret
Dim za1 As st_savret1
Dim nom As String
nom = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavret = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1

End Sub
Public Sub recopie_stock0(ByVal nom1 As String)
Dim za As st_savstock
Dim za1 As st_savsto1
Dim nom As String, nom2 As String
nom = chemin_app + "etude.boh"
nom2 = chemin_app + "etude.boh"
   lhFicDbf = FreeFile
    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
        za1.stsavstock = za
        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Put #lhFicDbf1, FileLength, za1
    End If

Loop
    Close #lhFicDbf
    Close #lhFicDbf1
End Sub
Private Sub recopie_fich0()
Dim nom1 As String
    nom1 = chemin_app + "etude.boh"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    Call init_etude0
    nom1 = chemin_app + "bassin.bin"
    If Dir(nom1) <> "" Then
        Call recopie_bv0(nom1)
    End If
    nom1 = chemin_app + "ouvrages.bin"
    If Dir(nom1) <> "" Then
        Call recopie_chute0(nom1)
    End If
    nom1 = chemin_app + "conduites.bin"
    If Dir(nom1) <> "" Then
        Call recopie_conduite0(nom1)
    End If
    nom1 = chemin_app + "ouvrages1.bin"
    If Dir(nom1) <> "" Then
        Call recopie_decant0(nom1)
    End If
    nom1 = chemin_app + "deversoir.bin"
    If Dir(nom1) <> "" Then
        Call recopie_deversoir0(nom1)
    End If
    nom1 = chemin_app + "retention.bin"
    If Dir(nom1) <> "" Then
        Call recopie_ret0(nom1)
    End If
    nom1 = chemin_app + "siphon.bin"
    If Dir(nom1) <> "" Then
        Call recopie_siphon0(nom1)
    End If
    nom1 = chemin_app + "stockage.bin"
    If Dir(nom1) <> "" Then
        Call recopie_stock0(nom1)
    End If
    nom1 = chemin_app + "bassin.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "ouvrages.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "conduites.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "ouvrages1.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "deversoir.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "retention.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "siphon.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
    nom1 = chemin_app + "stockage.bin"
    If Dir(nom1) <> "" Then
        Kill nom1
    End If
End Sub

