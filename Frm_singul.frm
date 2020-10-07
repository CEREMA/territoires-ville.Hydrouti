VERSION 5.00
Begin VB.Form Frm_singul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saisie des coudes"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "Frm_singul.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6585
   Begin VB.CommandButton Cmd_Quit 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   255
      Left            =   4560
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   3120
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   9
      Left            =   4800
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4230
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   3870
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2790
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   9
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   32
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   8
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   29
      Top             =   3840
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   7
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   26
      Top             =   3480
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   6
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   23
      Top             =   3120
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   20
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   17
      Top             =   2400
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   3
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   14
      Top             =   2040
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   11
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   9
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   8
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   28
      Top             =   3840
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   7
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   25
      Top             =   3480
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   6
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   22
      Top             =   3120
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   19
      Top             =   2760
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   16
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   3
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2040
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1680
      Width           =   850
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1320
      Width           =   850
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   9
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   30
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   8
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   27
      Top             =   3840
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   7
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   24
      Top             =   3480
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   6
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   21
      Top             =   3120
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2760
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2400
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   3
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2040
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1320
      Width           =   500
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   9
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   8
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3840
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   7
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3480
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   6
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3120
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   5
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2760
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   4
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   3
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2040
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1680
      Width           =   850
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1320
      Width           =   850
   End
   Begin VB.CommandButton Cmd_Del 
      Caption         =   "Supprimer"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   990
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ray 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   5
      Top             =   960
      Width           =   1000
   End
   Begin VB.TextBox Tb_Ang 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   4
      Top             =   960
      Width           =   850
   End
   Begin VB.TextBox Tb_Nbre 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   3
      Top             =   960
      Width           =   500
   End
   Begin VB.TextBox Tb_Type 
      BackColor       =   &H80000016&
      Height          =   300
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   850
   End
   Begin VB.CommandButton Cmd_aj_vif 
      Caption         =   "Ajout coude angle vif"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Cmd_aj_arr 
      Caption         =   "Ajout coude arrondi"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Lb_Ray 
      Alignment       =   2  'Center
      Caption         =   "Rayon mm"
      Height          =   255
      Left            =   3600
      TabIndex        =   37
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Lb_Ang 
      Alignment       =   2  'Center
      Caption         =   "Angle °"
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Lb_Nbre 
      Alignment       =   2  'Center
      Caption         =   "Nbre"
      Height          =   255
      Left            =   1800
      TabIndex        =   35
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Lb_Type 
      Alignment       =   2  'Center
      Caption         =   "Type"
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Frm_singul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Private Listcoud_anc As st_listcoude
Private ok_singul As Boolean
Private list_tb() As Variant
Private sval_champ As String
Private iSels As Integer
Private iSell As Integer
Private bKP As Boolean

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
Private Sub Cmd_aj_arr_Click()
Dim i As Integer, reponse As Integer
    i = rech_sui()
    If i > 0 Then
        Me.Tb_Type(i - 1).Visible = True
        Me.Tb_Nbre(i - 1).Visible = True
        Me.Tb_Ang(i - 1).Visible = True
        Me.Tb_Ray(i - 1).Visible = True
        Me.Cmd_Del(i - 1).Visible = True
        Me.Tb_Type(i - 1).Text = "Arrondi"
        Me.Tb_Nbre(i - 1).Text = "1"
        Me.Tb_Ang(i - 1).Text = "0.0"
        Me.Tb_Ray(i - 1).Text = "0.0"
   Else
        reponse = MsgBox("Le nombre de coudes est limité à 10", , "Saisie des coudes")
    End If
End Sub
Private Function rech_sui() As Integer
Dim i As Integer, j As Integer
j = 0
For i = 1 To 10
    If Trim(Me.Tb_Type(i - 1)) = "" Then
        j = i
        i = 10
    End If
Next
rech_sui = j
End Function

Private Sub Cmd_aj_vif_Click()
Dim i As Integer, reponse As Integer
    i = rech_sui()
    If i > 0 Then
        Me.Tb_Type(i - 1).Visible = True
        Me.Tb_Nbre(i - 1).Visible = True
        Me.Tb_Ang(i - 1).Visible = True
        Me.Tb_Ray(i - 1).Visible = False
        Me.Cmd_Del(i - 1).Visible = True
        Me.Tb_Type(i - 1).Text = "Angle vif"
        Me.Tb_Nbre(i - 1).Text = "1"
        Me.Tb_Ang(i - 1).Text = "0.0"
        Me.Tb_Ray(i - 1).Text = "0.0"
   Else
        reponse = MsgBox("Le nombre de coudes est limité à 10", , "Saisie des coudes")
    End If
End Sub




Private Sub Cmd_Del_Click(Index As Integer)

    Me.Tb_Type(Index).Visible = False
    Me.Tb_Nbre(Index).Visible = False
    Me.Tb_Ang(Index).Visible = False
    Me.Tb_Ray(Index).Visible = False
    Me.Cmd_Del(Index).Visible = False
    Me.Tb_Type(Index).Text = ""
    Me.Tb_Nbre(Index).Text = "0"
    Me.Tb_Ang(Index).Text = "0.0"
    Me.Tb_Ray(Index).Text = "0.0"
    If Index < 9 Then
        For i = Index To 8
            Me.Tb_Type(i).Visible = Me.Tb_Type(i + 1).Visible
            Me.Tb_Nbre(i).Visible = Me.Tb_Nbre(i + 1).Visible
            Me.Tb_Ang(i).Visible = Me.Tb_Ang(i + 1).Visible
            Me.Tb_Ray(i).Visible = Me.Tb_Ray(i + 1).Visible
            Me.Cmd_Del(i).Visible = Me.Cmd_Del(i + 1).Visible
            Me.Tb_Type(i).Text = Me.Tb_Type(i + 1).Text
            Me.Tb_Nbre(i).Text = Me.Tb_Nbre(i + 1).Text
            Me.Tb_Ang(i).Text = Me.Tb_Ang(i + 1).Text
            Me.Tb_Ray(i).Text = Me.Tb_Ray(i + 1).Text
            Me.Tb_Type(9).Visible = False
            Me.Tb_Nbre(9).Visible = False
            Me.Tb_Ang(9).Visible = False
            Me.Tb_Ray(9).Visible = False
            Me.Cmd_Del(9).Visible = False
            Me.Tb_Type(9).Text = ""
            Me.Tb_Nbre(9).Text = "0"
            Me.Tb_Ang(9).Text = "0.0"
            Me.Tb_Ray(9).Text = "0"
        Next
    End If
End Sub

Private Sub Cmd_ok_Click()
Dim i As Integer
For i = 1 To 10
    Listcoud.coude(i - 1).type = Tb_Type(i - 1).Text
    Listcoud.coude(i - 1).Nbre = txtVersNum(Tb_Nbre(i - 1).Text)
    Listcoud.coude(i - 1).angle = Round(txtVersNum(Tb_Ang(i - 1).Text), 2)
'    Listcoud.coude(i - 1).Rayon = Round(txtVersNum(Tb_Ray(i - 1).Text), 2)
' modification unité saisie  rayon mmm     'verification saisie entière?
    Listcoud.coude(i - 1).Rayon = Round(txtVersNum(Tb_Ray(i - 1).Text), 0) / 1000# ' en cas de rayon en mm (attention a l'arrondi)
Next
ok_singul = True
Unload Me
End Sub

Private Sub Cmd_Quit_Click()
Unload Me
End Sub


Private Sub Form_Load()
Centre Me
Set owner = MDIFrm_menu.rec_owner
    bKP = False
    sval_champ = ""

ok_singul = False
Listcoud_anc = Listcoud
Call ini_frm_singul
 Call init_l_tab
 Call donne_focus(Me)

End Sub
Private Sub ini_frm_singul()
Dim i As Integer
For i = 1 To 10
    If Trim(Listcoud_anc.coude(i - 1).type) <> "" Then
        Me.Tb_Type(i - 1).Visible = True
        Me.Tb_Type(i - 1).Enabled = False
        Me.Tb_Nbre(i - 1).Visible = True
        Me.Tb_Ang(i - 1).Visible = True
        Me.Cmd_Del(i - 1).Visible = True
        If Trim(Listcoud_anc.coude(i - 1).type) = "Arrondi" Then
            Me.Tb_Ray(i - 1).Visible = True
        Else
            Me.Tb_Ray(i - 1).Visible = False
        End If
    Else
        Me.Tb_Type(i - 1).Visible = False
        Me.Tb_Type(i - 1).Enabled = False
        Me.Tb_Nbre(i - 1).Visible = False
        Me.Tb_Ang(i - 1).Visible = False
        Me.Tb_Ray(i - 1).Visible = False
        Me.Cmd_Del(i - 1).Visible = False
    End If
        Me.Tb_Type(i - 1).Text = Listcoud_anc.coude(i - 1).type
        Me.Tb_Nbre(i - 1).Text = rempl_virgule(Format(Listcoud_anc.coude(i - 1).Nbre, "##0"))
        Me.Tb_Ang(i - 1).Text = rempl_virgule(Format(Listcoud_anc.coude(i - 1).angle, "##0.00"))
'        Me.Tb_Ray(i - 1).Text = rempl_virgule(Format(Listcoud_anc.coude(i - 1).Rayon, "#####0.00"))
' modif si saisie rayon en mm ???? ne pas oublier l'unité ds le titre
        Me.Tb_Ray(i - 1).Text = rempl_virgule(Format(Listcoud_anc.coude(i - 1).Rayon * 1000, "#######0"))
Next
End Sub

Private Sub Lb_Ray_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_Ray(Index).Text) <= Tb_Ray(Index).MaxLength Then
    KeyAscii = verif_car(Tb_Ray(Index).Text, KeyAscii, "Saisie Rayon", "R")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    owner.fobjet.enabled = True
    If ok_singul Then
         ebsiphon.List_coude = Listcoud
         If ebsiphon.ds > 0 Then
            owner.calc_kc
'            Call Frm_siphon.calc_kc
         End If
    End If
End Sub

Private Sub Tb_Ang_Change(Index As Integer)
Dim nom As String

If bKP Then
         nom = verif_cart0(Tb_Ang(Index).Text, "Saisie angle", "R")
  If nom = "" Then
    Tb_Ang(Index).Text = sval_champ
    Tb_Ang(Index).SelStart = iSels
    Tb_Ang(Index).SelLength = iSell
  End If
End If
'****

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Ang_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_Ang(Index).Text
    iSels = Tb_Ang(Index).SelStart
    iSell = Tb_Ang(Index).SelLength
    bKP = True
'   If Len(Tb_Ang(Index).Text) <= Tb_Ang(Index).MaxLength Then
'        KeyAscii = verif_car(Tb_Ang(Index).Text, KeyAscii, "Saisie angle", "R")
'    End If
End If
End Sub

Private Sub Tb_Nbre_Change(Index As Integer)
Dim nom As String

If bKP Then
        nom = verif_cart0(Tb_Nbre(Index).Text, "Saisie nombre", "I")
  If nom = "" Then
    Tb_Nbre(Index).Text = sval_champ
    Tb_Nbre(Index).SelStart = iSels
    Tb_Nbre(Index).SelLength = iSell
  End If
End If
'****

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Nbre_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_Nbre(Index).Text
    iSels = Tb_Nbre(Index).SelStart
    iSell = Tb_Nbre(Index).SelLength
    bKP = True
'    If Len(Tb_Nbre(Index).Text) <= Tb_Nbre(Index).MaxLength Then
'        KeyAscii = verif_car(Tb_Nbre(Index).Text, KeyAscii, "Saisie nombre", "I")
'    End If
End If
End Sub


Private Sub Tb_Ray_Change(Index As Integer)
Dim nom As String

If bKP Then
         nom = verif_cart0(Tb_Ray(Index).Text, "Saisie rayon", "I")
  If nom = "" Then
    Tb_Ray(Index).Text = sval_champ
    Tb_Ray(Index).SelStart = iSels
    Tb_Ray(Index).SelLength = iSell
  End If
End If
'****

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Ray_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_Ray(Index).Text
    iSels = Tb_Ray(Index).SelStart
    iSell = Tb_Ray(Index).SelLength
    bKP = True
'   If Len(Tb_Ray(Index).Text) <= Tb_Ray(Index).MaxLength Then
'        KeyAscii = verif_car(Tb_Ray(Index).Text, KeyAscii, "Saisie rayon", "R")
'    End If
End If

End Sub
