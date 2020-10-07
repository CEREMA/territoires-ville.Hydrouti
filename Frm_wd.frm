VERSION 5.00
Begin VB.Form Frm_wd 
   Caption         =   "Ouverture WORD"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Frm_wd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu
Public Type1 As String
Public titre1 As String
Public sstitre1 As String
Public ssTitre2 As String
Public ssTitre3 As String
Public ssTitre4 As String
Public ssTitre5 As String
Public awd As Word.Application
Public nom_fic As String
Public ad As Word.Document

Private Sub Form_Load()
Dim wrstyles As Variant
Dim npar As Integer, i As Integer
Dim stylew As Style, stylew1 As Style, stylew2 As Style, stylew3 As Style
Dim stylew4 As Style
Set owner = MDIFrm_menu.rec_owner

If awd Is Nothing Then
    Set awd = New Word.Application
Else
    If awd.Documents.Count > 0 Then
        If Not ad Is Nothing Then
            Debug.Print ad.Name
             ad.Close
        End If
    End If
    Set ad = Nothing
End If
If Dir(nom_fic) <> "" Then
Set ad = awd.Documents.Open(FileName:=nom_fic)
npar = ad.Paragraphs.Count - 1
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, _
    End:=ad.Paragraphs(npar).Range.End)
arange.InsertParagraphAfter
npar = npar + 1
Else
Set ad = awd.Documents.Add
npar = 0
End If
wrstyles = awd.Languages(wdFrench).WritingStyleList
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
ad.Styles(73).Font.Size = 22
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

      
'cframes.Add Item:=Me.Frame1, Key:=CStr(1)

ad.Range.InsertAfter "BOHHA"
npar = npar + 1
npard = npar
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew1)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
Set f = ad.Frames.Add(Range:=ad.Paragraphs(npar).Range) 'ajouter un cadre autour d'un paragraphe
With f.Shading
    .Texture = wdTexture20Percent 'fond en gris
End With
    f.Borders.Enable = False
   
'    f.Width = 120
'ad.Paragraphs(npar).Range.f.TextWrap = True
'ad.Range.InsertParagraphAfter
'npar = npar + 1

ad.Range.InsertAfter "Boite à Outils Hydrologie , Hydraulique et Assainissement"
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew4)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
ad.Range.InsertAfter "Centre d' Etudes Techniques de l' Equipement de l'Est"
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
ad.Range.InsertAfter "Laboratoire Régional de Nancy"
ad.Range.InsertParagraphAfter
npar = npar + 1
nparf = npar
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
'ad.Paragraphs(npar).Range.Font.Underline = wdUnderlineDouble
Set arange = ad.Range(Start:=ad.Paragraphs(npard).Range.Start, _
    End:=ad.Paragraphs(nparf).Range.End)
Set f = ad.Frames.Add(arange) 'ajouter un cadre autour d'un paragraphe



ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Range.InsertParagraphAfter
npar = npar + 1
npard = npar
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Range.InsertAfter Me.titre1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew3)
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
nparf = npar
Set arange = ad.Range(Start:=ad.Paragraphs(npard).Range.Start, _
    End:=ad.Paragraphs(nparf).Range.End)
Set f = ad.Frames.Add(arange) 'ajouter un cadre autour d'un paragraphe
With f.Shading
    .Texture = wdTexture20Percent 'fond en gris
End With
    f.Borders.Enable = False
Select Case Type1
    Case Is = "decant"
        For j = 1 To 2
            ad.Range.InsertParagraphAfter
            npar = npar + 1
            ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        Next
'******liste 1***************************************************
        npar = cre_table(npar, sstitre1, "list_don1", stylew2, wrstyles)
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
'******liste 2***************************************************
        npar = cre_table(npar, ssTitre2, "list_int1", stylew2, stylew)
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
'******liste 3***************************************************
        npar = cre_table(npar, ssTitre3, "list_resu1", stylew2, stylew)
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
'*******************************************************
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        npar = recup_dess(npar, 452, 180, "dess.bmp")
     Case Is = "stockage"
        For j = 1 To 1
            ad.Range.InsertParagraphAfter
            npar = npar + 1
            ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        Next
'******liste 1***************************************************
        npar = cre_table(npar, sstitre1, "list_don1", stylew2, wrstyles)
'******liste 2***************************************************
        npar = cre_table(npar, ssTitre2, "list_int1", stylew2, stylew)
'******liste 3***************************************************
        npar = cre_table(npar, ssTitre3, "list_resu1", stylew2, stylew)
'*******************************************************
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        npar = recup_dess(npar, 452, 180, "dess.bmp")
End Select
'***********************************************

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
nom = "c:\hydraulique\bo_v4\essai00.doc"
DoEvents
awd.Application.Visible = True
DoEvents
ad.SaveAs nom

'Ad.Close
'A.Application.Qu

End Sub
Private Function cre_table(ByVal npar As Integer, ByVal Titre As String, ByVal nomfich As String, _
    ByVal style2 As Style, ByRef wrstyles As Variant) As Integer
Dim liste() As Variant
Dim nb As Integer, i As Integer
Dim ct As Table
liste = owner.fobjet.lect_list(nomfich)
nb = UBound(liste)
ad.Range.InsertAfter Titre
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(style2)
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, End:=ad.Paragraphs(npar).Range.End)
'Set ct = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=3)
'************création du tableau
Set ct = ad.Tables.Add(Range:=arange, NumRows:=nb + 1, NumColumns:=4)
    ct.Columns(1).SetWidth ColumnWidth:=CentimetersToPoints(1.5), _
        RulerStyle:=wdAdjustProportional
    ct.Columns(2).SetWidth ColumnWidth:=CentimetersToPoints(9#), _
        RulerStyle:=wdAdjustProportional
    ct.Columns(3).SetWidth ColumnWidth:=CentimetersToPoints(4), _
        RulerStyle:=wdAdjustProportional
'    ct.Columns(4).SetWidth ColumnWidth:=CentimetersToPoints(2), _
'       RulerStyle:=wdAdjustProportional
    ct.Borders(wdBorderHorizontal) = False
    ct.Borders(wdBorderVertical) = False
'*******************************
       For i = 1 To nb + 1
             ct.Cell(i, 2).Select
             Selection.TypeText Text:=liste(i - 1, 1)
             ct.Cell(i, 3).Select
             Selection.TypeText Text:=liste(i - 1, 2)
             Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
             ct.Cell(i, 4).Select
             Selection.TypeText Text:=liste(i - 1, 3)
        Next
    npar = ad.Paragraphs.Count - 1
    cre_table = npar
End Function
Private Function recup_dess(ByVal npar As Integer, ByVal xlar As String, ByVal haut As String, _
    ByVal nomfich As String) As Integer
'**cree un cadre pour insérer un dessin*****
        Set maplage = ad.Range(Start:=ad.Paragraphs(npar).Range.End, _
        End:=ad.Paragraphs(npar).Range.End)
        Set f = ad.Frames.Add(Range:=maplage)
        f.Width = xlar
        f.Height = haut
        With f.Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble 'Options.DefaultBorderLineStyle
    '        .LineWidth = Options.DefaultBorderLineWidth
    '        .ColorIndex = Options.DefaultBorderColorIndex
        End With
        With f.Borders(wdBorderLeft)
            .LineStyle = wdLineStyleDouble
        End With
        With f.Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
        End With
        With f.Borders(wdBorderRight)
            .LineStyle = wdLineStyleDouble
        End With
'        ad.Range.InsertParagraphAfter
'        npar = npar + 1
'*** insertion du dessin*****
            maplage.InlineShapes.AddPicture FileName:=chemin_app + nomfich, _
                 LinkToFile:=False, SaveWithDocument:=True
    recup_dess = npar
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not awd Is Nothing Then
    ad.Close
    awd.Quit
   Set ad = Nothing
   Set awd = Nothing
End If
End Sub

