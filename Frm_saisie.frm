VERSION 5.00
Begin VB.Form Frm_saisie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saisie d'informations"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "Frm_saisie.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8790
   Begin VB.CommandButton cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   300
      Left            =   6480
      TabIndex        =   4
      Top             =   2280
      Width           =   1250
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   4800
      TabIndex        =   3
      Top             =   2280
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   7695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      MaxLength       =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Saisie du service (2 lignes 60 caractères maxi)"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Frm_saisie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nom_fichier_texte As String
Private owner As MDIFrm_menu

Private Sub Cmd_annul_Click()
    Unload Me
End Sub

Private Sub Cmd_ok_Click()
Dim esave As st_texte
    text_serv1 = Trim(Text1.Text)
    text_serv2 = Trim(Text2.Text)
 'mise à jour REGEDIT   /HKEY_CURRENT_USER/Software/VB and VBA..../
    SaveSetting "Hydrouti", "Informations", "Info1", text_serv1
    SaveSetting "Hydrouti", "Informations", "Info2", text_serv2
'    lhFicDbf = FreeFile
'    Open nom_fichier_texte For Random Access Read Write As #lhFicDbf Len = Len(esave)
''        FileLength = LOF(lhFicDbf) / Len(esave) + 1
'    Put #lhFicDbf, 1, esave
'    Close #lhFicDbf
    Unload Me
End Sub

Private Sub Form_Load()
    Centre Me
    Set owner = MDIFrm_menu.rec_owner
 '   nom_fichier_texte = chemin_app + "service.bin"
    Text1.Text = Trim(text_serv1)
    Text2.Text = Trim(text_serv2)
End Sub

