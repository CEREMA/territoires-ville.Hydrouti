VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sauvegarde"
   ClientHeight    =   1635
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   4300
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "OK"
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1250
   End
   Begin VB.CommandButton Cmd_annul 
      Caption         =   "Annuler"
      Height          =   300
      Left            =   3240
      TabIndex        =   0
      Top             =   1200
      Width           =   1250
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private owner As MDIFrm_menu
Private Sub Cmd_annul_Click()
    Unload Me
End Sub

Private Sub Cmd_ok_Click()
If Trim(Text1.Text) <> "" Then
    If Not owner.fbassin Is Nothing Then
        owner.fbassin.Tb_titre = Text1.Text
        Call owner.fbassin.save
    End If
    If Not owner.fobjet Is Nothing Then
        owner.fobjet.Tb_titre = Text1.Text
        Call owner.fobjet.save(True)
    End If
    Unload Me
Else
    reponse = MsgBox("Le nom du siphon n'est pas renseigné.", , "Sauvegarde d'un siphon")
End If
End Sub

Private Sub Form_Load()
    Set owner = MDIFrm_menu.rec_owner
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not owner.fbassin Is Nothing Then
        owner.fbassin.Enabled = True
    End If
    If Not owner.fobjet Is Nothing Then
        owner.fobjet.Enabled = True
    End If
End Sub

