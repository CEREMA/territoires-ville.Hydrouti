VERSION 5.00
Begin VB.Form Frm_savword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sauvegarde sous WORD"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "Frm_savword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Frm_savword.frx":08CA
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_complet 
      Caption         =   "Compléter"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_rempl 
      Caption         =   " Remplacer"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   1000
   End
End
Attribute VB_Name = "Frm_savword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sav_w As Boolean
Public mod_sav As String

Private Sub Cmd_annul_Click()
    Me.sav_w = True
    Me.mod_sav = ""
    Unload Me
End Sub

Private Sub Cmd_complet_Click()
    Me.sav_w = False
    Me.mod_sav = "complete"
    Unload Me
End Sub

Private Sub Cmd_rempl_Click()
'    Kill (Frm_imp.Label1.Caption)
    Me.sav_w = False
    Me.mod_sav = "remplace"
    Unload Me
End Sub

Private Sub Form_Load()
    Centre Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Frm_imp.Enabled = True
End Sub
