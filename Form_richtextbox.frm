VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   105.04
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   146.05
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   4800
      Width           =   6015
   End
   Begin VB.CommandButton Cmd_word 
      Caption         =   "WORD"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin RichTextLib.RichTextBox Rtb1 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form_richtextbox.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Public Titre1 As String
Public ssTitre1 As String
Public ssTitre2 As String
Public ssTitre3 As String
Public ssTitre4 As String
Public ssTitre5 As String

Private Sub Cmd_word_Click()
Dim liste() As Variant
Dim nb As Integer
Dim nb_ordre As Integer
Dim xdate As String, chaine As String, nomtable As String, sreq As String
Dim nombase As String, nom As String
Dim a As Word.Application
Dim wrstyles As Variant
Dim npar As Integer, i As Integer, isel As Integer, icont As Integer
Dim j As Integer
Dim exist As Boolean, ok As Boolean, ok1 As Boolean
Dim stylew As Object, stylew1 As Object, stylew2 As Object, stylew3 As Object
'Dim cframes As New Collection
'        If Not open_dbsuivi Then
'            nombase = chemin_suivi + "suivi.mdb"
'            Set dbsuivi = OpenDatabase(nombase)
'       End If
npar = 0

'Set a = CreateObject("WORD.DOCUMENT.8")
' modification julienne
Set a = New Word.Application
Set ad = a.Documents.Add
wrstyles = a.Languages(wdFrench).WritingStyleList
'ad.Shapes.SelectAll
'Debug.Print ad.Shapes.Count
'ad.shapes.AddShape msoShapeRectangle, 5, 5, 10, 20
'ad.shapes.AddShape 96, 5, 5, 10, 20


'With ad.Shapes.AddShape(msoShapeRectangle, 0, 0, 5, 4).Fill
'    .ForeColor.RGB = RGB(128, 0, 0)
'    .BackColor.RGB = RGB(170, 170, 170)
'    .TwoColorGradient msoGradientHorizontal, 1
'End With
'styles utilisés
 Set stylew1 = ad.Styles(71)
ad.Styles(71).Font.Size = 25
ad.Styles(71).Font.Bold = True
ad.Styles(71).Font.Italic = False
 Set stylew2 = ad.Styles(72)
ad.Styles(72).Font.Size = 11
ad.Styles(72).Font.Bold = True
ad.Styles(72).Font.Italic = False
Set stylew3 = ad.Styles(73)
ad.Styles(73).Font.Size = 15
ad.Styles(73).Font.Bold = True
ad.Styles(73).Font.Italic = False
Set stylew4 = ad.Styles(73)
ad.Styles(73).Font.Size = 12
ad.Styles(73).Font.Bold = True
ad.Styles(73).Font.Italic = False
Set stylew = ad.Styles(49)
ad.Styles(49).Font.Size = 10
ad.Styles(49).Font.Bold = False
ad.Styles(49).Font.Italic = False
'fin styles
'For j = 1 To 6
'    Ad.Range.InsertParagraphAfter
'    npar = npar + 1
'Next

'  Rtb1.SelText = "BOHHA  "

      
'cframes.Add Item:=Me.Frame1, Key:=CStr(1)
ad.Range.InsertAfter "BOHHA"
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew1)
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set f = ad.Frames.Add(Range:=ad.Paragraphs(npar).Range) 'ajouter un cadre autour d'un paragraphe
With f.Shading
    .Texture = wdTexture20Percent 'fond en gris
End With
    f.Borders.Enable = False
'    f.Width = 120
'ad.Paragraphs(npar).Range.f.TextWrap = True

ad.Range.InsertAfter "Boite à Outils Hydrologie , Hydraulique et Assainissement"
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew4)
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
'ad.Paragraphs(npar).Range.Font.Underline = wdUnderlineDouble



ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Range.InsertAfter Titre1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew3)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
Set f = ad.Frames.Add(Range:=ad.Paragraphs(npar).Range) 'ajouter un cadre autour d'un paragraphe
With f.Shading
    .Texture = wdTexture20Percent 'fond en gris
End With
    f.Borders.Enable = False
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Range.InsertParagraphAfter
npar = npar + 1
'For j = 1 To 9
'    Ad.Range.InsertParagraphAfter
'    npar = npar + 1
'Next
ad.Range.InsertParagraphAfter
npar = npar + 1
    liste = owner.fobjet.lect_list("list_don1")
    nb = UBound(liste)
ad.Range.InsertAfter ssTitre1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew2)
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
ad.Range.InsertParagraphAfter
npar = npar + 1
For i = 0 To nb
'    icont = 0
'        If icont > 48 Then
'            Set myRange = Ad.Range
'            With myRange
'                .Collapse Direction:=wdCollapseEnd
'                .InsertBreak type:=wdPageBreak
'            End With
'            Ad.Range.InsertParagraphAfter
'            npar = npar + 1
'            icont = 1
'        End If
        chaine = liste(i, 1)
        ad.Range.InsertAfter chaine
        ad.Range.InsertParagraphAfter
        npar = npar + 1
'        icont = icont + 1
Next
                            ad.Paragraphs(npar).Range.Font.Bold = True


'Cet exemple montre comment ajouter un modèle de trait à myDocument.

'With ad.Shapes.AddLine(10, 100, 250, 0).Line
'    .Weight = 6
'    .ForeColor.RGB = RGB(0, 0, 255)
'    .BackColor.RGB = RGB(128, 0, 0)
'    .Pattern = msoPatternDarkDownwardDiagonal
'End With
'Ad.Sections(1).Footers(wdHeaderFooterPrimary).Range _
' .InsertDateTime DateTimeFormat:="jj MMMM aaaa", _
' InsertAsField:=True
'With Ad.Sections(1).Footers(wdHeaderFooterPrimary)
'    .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight
'End With
'With Ad.Sections(1).Headers(wdHeaderFooterPrimary)
'    .Range.InsertAfter nom_rapport
'    .Range.Paragraphs.Alignment = wdAlignParagraphCenter
'    .Range.Paragraphs.Borders.Enable = True
'End With
'If nom_rapport = "" Then
'    nom = "c:\rapport.doc"
'Else
'    nom = nom_rapport + ".doc"
'End If
'del_fic (nom)
nom = "c:\hydraulique\bo_v4\essai00.doc"
DoEvents
a.Application.Visible = True
DoEvents
ad.SaveAs nom

'Ad.Close
'A.Application.Quit
Set a = Nothing
'If Not open_dbsuivi Then
'      dbsuivi.Close
'End If
End Sub



Private Sub Command1_Click()
Dim liste() As Variant
Dim nb As Integer
'Rtb1.BackColor = couleur.blanc
'  Me.WindowState = 2 'plein ecran
Rtb1.BorderStyle = rtfFixedSingle
Rtb1.Top = 20
Rtb1.Left = 30
Rtb1.Height = 297
Rtb1.Width = 210
'Rtb1.Font.Italic = True
'Rtb1.Font.Italic = True
'Rtb1.Font.Size = 12
  Rtb1.SelFontSize = 40
  Rtb1.SelIndent = 0
  Rtb1.SelText = "BOHHA  "
'Rtb1.RightMargin = 3000
  Rtb1.SelFontSize = 15
'  Rtb1.SelIndent = 50
  Rtb1.SelText = " Boite à Outils Hydrologie,Hydraulique et Assainissement" + Chr(10) + Chr(13)
  Rtb1.SelIndent = 20
  Rtb1.SelFontSize = 18
Rtb1.SelText = Titre1
'Rtb1.SelStart = 5
'Rtb1.SelLength = 10
'Debug.Print Rtb1.SelText
'Rtb1.SelBold = True
'Rtb1.SelColor = 255
'Rtb1.SelBold


'''Rtb1.RightMargin = 3000
'''
'''  With Rtb1
'''      .SelTabCount = 5
'''      For X = 0 To .SelTabCount - 1
'''         .SelTabs(X) = 5 * X
'''      Next X
'''   End With
'''Rtb1.SelStart = Len(Rtb1.Text)
'''
'''Rtb1.SelText = "Bonjour" + Chr(10) ' + Chr(13)
'''Rtb1.SelStart = Len(Rtb1.Text)
'''Rtb1.SelText = "fffffffff"
'''Rtb1.SelStart = Len(Rtb1.Text)
'''Rtb1.SelText = Chr(9)
'''Rtb1.SelStart = Len(Rtb1.Text)
'''Rtb1.SelText = "fffffffffffffffff"
'''Rtb1.SelStart = Len(Rtb1.Text)
'''Rtb1.SelText = Chr(9) + "Bonjour"
'''Rtb1.SelStart = Len(Rtb1.Text)
'''Rtb1.SelText = Chr(9) + "Bonjour"

Rtb1.Container.Line (10, 10)-(120, 200), 255
   Rtb1.SelText = Chr(10) ' + Chr(13)
  Rtb1.SelIndent = 10
  Rtb1.SelFontSize = 12
  Rtb1.SelText = ssTitre1 + Chr(10) '+ Chr(13)

    liste = owner.fobjet.lect_list("list_don1")
Rtb1.SelStart = Len(Rtb1.Text)

    nb = UBound(liste)
  Rtb1.SelIndent = 20
  Rtb1.SelFontSize = 10
  With Rtb1
      .SelTabCount = 2
        .SelTabs(0) = 150
        .SelTabs(1) = 60
   End With
   For i = 0 To nb
        Rtb1.SelText = liste(i, 1)
        Rtb1.SelText = Chr(9) + liste(i, 2)
        Rtb1.SelText = Chr(9) + liste(i, 3) + Chr(10) ' + Chr(13)
        Rtb1.SelStart = Len(Rtb1.Text)
    Next
'        ytop = ecr_frm("Frm_par1", ssTitre1, "list_don1", ytop)
'        ytop = ytop + ydec
'        ytop = ecr_frm("Frm_par2", ssTitre2, "list_int1", ytop)
'        ytop = ytop + ydec
'        ytop = ecr_frm("Frm_par3", ssTitre3, "list_resu1", ytop)
'
End Sub

Private Sub Command2_Click()
Rtb1.SaveFile "c:\hydraulique\bo_v4\essai1.rtf"

Open "c:\hydraulique\bo_v4\essai2.rtf" For Output As 1
   
   Print #1, Rtb1.TextRTF
Close 1
End Sub

Private Sub Command3_Click()
Rtb1.LoadFile "c:\hydraulique\bo_v4\essai2.rtf"
End Sub

Private Sub Form_Load()
'Document1.Object "c:\hydraulique\bo_v4\essai.doc"
    Set owner = MDIFrm_menu.rec_owner
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode, Shift
End Sub

