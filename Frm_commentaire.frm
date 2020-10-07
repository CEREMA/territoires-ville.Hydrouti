VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_commentaire 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Aide"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1770
   ControlBox      =   0   'False
   Icon            =   "Frm_commentaire.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser RTB_Aide 
      Height          =   5400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   9525
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox RTB_com 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Frm_commentaire.frx":08CA
   End
End
Attribute VB_Name = "Frm_commentaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Private nom_fichier As String
Private nom_fichier_exemple As String
Private nom_champ As String
Private start_fic As Integer, end_fic As Double


Private Sub Cmd_com_Exemple()
RTB_aide.Navigate (nom_fichier_exemple + "#" + "")

End Sub

Private Sub Cmd_Com_Retour()
RTB_aide.Navigate (nom_fichier + "#" + nom_champ)
'    RTB_aide.GoBack
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Cmd_Com_Retour
If KeyCode = 113 Then Cmd_com_Exemple
End Sub


Private Sub Form_Load()
     Set owner = MDIFrm_menu.rec_owner
   Me.Width = 4500 'owner.Width / 6# '2000
  nom_fichier = ""
  nom_champ = ""
   affich_aide owner.chemin_aide + IDhlpAideFichier, "", owner.chemin_aide + IDhlpAideExempleFichier ' "Structure générale"
End Sub
Private Sub retaille()
    Me.Left = 0
    Me.Top = 0
End Sub
Public Sub retailler()
Call retaille
End Sub
Private Sub change_text_size0(ByRef rtb As RichTextBox, ByVal coef As Double)
For i = 0 To Len(rtb.Text)
    rtb.SelStart = i
    rtb.SelLength = 1
    rtb.SelFontSize = rtb.SelFontSize * coef
    If coef <> 1 Then Exit Sub
Next
End Sub
Private Sub change_text_size(ByRef rtb As RichTextBox, ByVal coef As Double)
Dim i As Double, ideb As Double, ifin As Double, szp As Double, sz As Double
Dim Longueur As Double
If coef = 1 Then Exit Sub
i = 0
ideb = i
ifin = i
szp = 0
rtb.SelStart = i
rtb.SelLength = 1
szp = rtb.SelFontSize
While i <= Len(rtb.Text)
rtb.SelStart = i
rtb.SelLength = 1
sz = rtb.SelFontSize
If sz <> szp Then
    rtb.SelStart = ideb
    rtb.SelLength = i - ideb
    rtb.SelFontSize = szp * coef
    ideb = i
    ifin = i
    szp = sz
End If
i = i + 1
Wend
    rtb.SelStart = ideb
    rtb.SelLength = i - ideb - 1
    rtb.SelFontSize = szp * coef
    rtb.SelStart = 0
    rtb.SelLength = 0
End Sub

Public Sub affich_aide(ByVal nom As String, ByVal texte As String, ByVal nom1 As String)
Dim aa As String
' ca marche
'If nom = "" Then
'    RTB_aide.SelStart = end_fic
'    i = RTB_aide.Find("HIERARCHISATION", 0)
'    RTB_aide.SelLength = 0
nom_fichier_exemple = nom1

If texte = "aide" Then
        RTB_aide.LoadFile (nom)
        Call change_text_size(RTB_aide, owner.coef)

Else
    If nom <> nom_fichier Or texte <> nom_champ Then
'        RTB_aide.LoadFile (nom)
'        Call change_text_size(RTB_aide, owner.coef)
'        start_fic = 1
'        end_fic = Len(RTB_aide.Text)
''       aa = RTB_aide.SelText
''
'RTB_aide.SelStart = end_fic
'        i = RTB_aide.Find(texte, 0, , rtfMatchCase)
'        RTB_aide.SelLength = 0
        RTB_aide.Navigate (nom + "#" + texte)
        nom_fichier = nom
        nom_champ = texte
    End If
End If
End Sub
Private Sub Form_Resize()
'    RTB_Com.Width = Me.Width - 100
     RTB_aide.Width = Me.Width - 100
    RTB_aide.Height = Me.Height - 400 '- RTB_Com.Height
    Set owner = MDIFrm_menu.rec_owner
    owner.change_taille
  
End Sub

Private Sub mnufichier_Click()
Call owner.fobjet.MnuQuit_Click
End Sub

