VERSION 5.00
Begin VB.Form Frm_titre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sauvegarde"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "Frm_titre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   4600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   4605
   End
End
Attribute VB_Name = "Frm_titre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Private Sub Cmd_annul_Click()
    Unload Me
End Sub

Private Sub Cmd_ok_Click()
If Trim(Text1.Text) <> "" Then
    If Not owner.fbassin Is Nothing Then
        owner.fbassin.Tb_titre = Text1.Text
        Call owner.fbassin.save(True)
Else
'    End If
'    If Not owner.fobjet Is Nothing Then
        owner.fobjet.Tb_titre = Text1.Text
        Call owner.fobjet.save(True)
    End If
    Unload Me
Else
    reponse = MsgBox("Le nom  n'est pas renseigné.", , "Sauvegarde ")
End If
End Sub

Private Sub Form_Load()
    Centre Me
    Set owner = MDIFrm_menu.rec_owner
End Sub
