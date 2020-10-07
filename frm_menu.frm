VERSION 5.00
Begin VB.Form frm_menu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HYDRAULIQUES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOITE à OUTILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   3855
      Left            =   960
      Top             =   600
      Width           =   6615
   End
   Begin VB.Menu m_Bassin 
      Caption         =   "Bassin Versant"
   End
   Begin VB.Menu m_DO 
      Caption         =   "Déversoir d'Orage"
   End
   Begin VB.Menu m_Retenue 
      Caption         =   "Retenue"
      Begin VB.Menu m_Stockage 
         Caption         =   "Bassin de Stockage Restitution"
      End
      Begin VB.Menu m_Retention 
         Caption         =   "Bassin de Retention Pluviale"
      End
   End
   Begin VB.Menu m_autre 
      Caption         =   "Autre"
      Begin VB.Menu m_chute 
         Caption         =   "Chute"
      End
      Begin VB.Menu m_decantation 
         Caption         =   "Décantation"
      End
      Begin VB.Menu m_siphon 
         Caption         =   "Siphon"
      End
   End
   Begin VB.Menu m_fermer 
      Caption         =   "Fermer"
   End
End
Attribute VB_Name = "frm_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Centre Me
    Call ini_color
    chemin_app = App.Path + "\"
    long_enreg = 363
    ini_bv
    do_bv = False
    sto_bv = False
    ret_bv = False
End Sub

Private Sub m_bassin_Click()
    frm_menu.Enabled = False
    frm_bv2.Show
End Sub

Private Sub m_do_Click()
    frm_menu.Enabled = False
    Frm_do.Show
End Sub

Private Sub m_Retention_Click()
    frm_menu.Enabled = False
    Frm_ret.Show
End Sub

Private Sub m_siphon_Click()
    frm_menu.Enabled = False
    Frm_siphon.Show
End Sub
Private Sub m_chute_Click()
    frm_menu.Enabled = False
    Frm_chute.Show
End Sub


Private Sub m_stockage_Click()
    frm_menu.Enabled = False
    Frm_stock.Show
End Sub
Private Sub m_decantation_Click()
    frm_menu.Enabled = False
    Frm_decant.Show
End Sub

Private Sub m_fermer_Click()
    Unload Me
End Sub
