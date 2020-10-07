VERSION 5.00
Begin VB.Form Frm_frontal 
   Caption         =   "Hydraulique Déversoir Frontal"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10620
   Begin VB.Frame Frm_bassin 
      Caption         =   "Bassin "
      Height          =   1695
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3855
      Begin VB.TextBox Tb_Qrin 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox Tb_Qts 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   3
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Tb_Qpluie 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2205
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Lb_Qrin 
         Caption         =   "Débit de rinçage "
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   1125
         Width           =   1815
      End
      Begin VB.Label Lb_Qts 
         Caption         =   "Débit de temps sec "
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   765
         Width           =   1815
      End
      Begin VB.Label Lb_Qpluie 
         Caption         =   "Débit d'eau pluviale "
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   1815
      End
      Begin VB.Label Lb_u1 
         Caption         =   "l/s"
         Height          =   300
         Left            =   3285
         TabIndex        =   7
         Top             =   405
         Width           =   450
      End
      Begin VB.Label Lb_u2 
         Caption         =   "l/s"
         Height          =   300
         Left            =   3285
         TabIndex        =   6
         Top             =   765
         Width           =   450
      End
      Begin VB.Label Lb_u3 
         Caption         =   "l/s"
         Height          =   300
         Left            =   3285
         TabIndex        =   5
         Top             =   1125
         Width           =   450
      End
   End
   Begin VB.CommandButton Cmd_Sel_Bv 
      Caption         =   "Sélection d'un bassin"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuquit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "Frm_frontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()

End Sub
