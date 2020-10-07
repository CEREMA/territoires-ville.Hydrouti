VERSION 5.00
Begin VB.Form Frm_graph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "graphique"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "Frm_graph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin hydrouti.UC_graphique UC_graphique1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
   End
End
Attribute VB_Name = "Frm_graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_annul_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Centre Me
End Sub


