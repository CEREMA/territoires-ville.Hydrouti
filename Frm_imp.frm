VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_imp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impression des résultats"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "Frm_imp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Nom de la fiche"
      Height          =   1455
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   6495
      Begin VB.CommandButton Cmd_valid 
         Caption         =   "Valider"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton Cmd_repos 
         Caption         =   "Repositionner"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   360
         Width           =   1365
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin MSComDlg.CommonDialog Cdlg2 
      Left            =   1200
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Imprimante"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   6495
      Begin VB.CommandButton Cmd_config 
         Caption         =   "Configurer"
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   480
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   3840
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nom complet du fichier"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   320
         Width           =   4575
      End
      Begin VB.CommandButton Cmd_recfic 
         Caption         =   "Enregistrer sous.."
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "Aperçu"
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   3840
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      Begin VB.OptionButton Opt_OOO 
         Caption         =   "Fiche OpenOffice"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   1800
      End
      Begin VB.OptionButton Opt_word 
         Caption         =   "Fichier WORD"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Width           =   1800
      End
      Begin VB.OptionButton Opt_imp 
         Caption         =   "Imprimante"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.Label Label3 
      Height          =   135
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   6375
   End
End
Attribute VB_Name = "Frm_imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public owner As MDIFrm_menu
Public Type1 As String
Public nomobjet As String
Public titre1 As String
Public sstitre1 As String
Public ssTitre2 As String
Public ssTitre3 As String
Public ssTitre4 As String
Public ssTitre5 As String
Public ssTitre6 As String
Public des1_titrh As String
Public des1_titrb As String
Public des2_titrh As String
Public des2_titrb As String
Public nom_fic As String
Public nom_fiche As String
Private ad As Word.Document
Public stylew As Word.Style
Public stylew1 As Word.Style
Public stylew0 As Word.Style
Public stylew2 As Word.Style
Public stylew3 As Word.Style
Public stylew4 As Word.Style
Public wrstyles As Variant
Private sav_word As Boolean
Private mod_save As String
'Houpie 20040123 modif ajout messages
'Public fso As New FileSystemObject
'Public exportTxt As TextStream
'Public TotoTxt As String, MyAppli
''''''''''''''''''''
Private Sub Cmd_annul_Click()
    Unload Me
End Sub

Private Sub Cmd_config_Click()
On Error GoTo erreur:
Dim oImp As Printer
Printer.TrackDefault = True
cdlg1.PrinterDefault = True
'Cdlg1.Flags = cdlPDPrintSetup Or cdlPDReturnDefault
'Cdlg1.ShowPrinter
cdlg1.Flags = cdlPDPrintSetup ' Or cdlPDReturnDC  'Or cdlPDReturnDefault
cdlg1.CancelError = True
Dim snp As String
cdlg1.Orientation = cdlPortrait
cdlg1.ShowPrinter
Printer.Orientation = cdlPortrait
While Printer.Orientation = cdlLandscape
'If Cdlg1.Orientation = cdlLandscape Then
    MsgBox "l'impression doit se faire en mode portrait", vbExclamation, _
        "Configuration imprimante"
'    Cdlg1.Orientation = cdlPortrait
    cdlg1.CancelError = True
   cdlg1.ShowPrinter
'End If
Wend
'Dim X As Printer
'snp = Cdlg1.Tag
'For Each X In Printers
'For i = 0 To Printers.Count - 1
'   Debug.Print Printers(i).hDC, Cdlg1.hDC, Printers(i).TrackDefault
' '  If Printers(i).hDC = Cdlg1.hDC Then
'   If Printers(i).DeviceName = "HP LaserJet 4/4M" Then
'      ' Définit l'imprimante comme imprimante par
'      ' défaut du système.
'      Set Printer = Printers(i)
'      ' Cesse la recherche d'imprimante.
'      Exit For
'   End If
'Next

Label2.Caption = Printer.DeviceName
Exit Sub
erreur:
'bannul = True
Resume Next
End Sub

Private Sub Cmd_ok_Click()
Dim reponse As Integer
Dim message As String
If Opt_imp Then
    Unload Me
    FrmPrint.Show
ElseIf Opt_word Then
   Label3.Caption = ""
   Me.Cmd_ok.Enabled = False
   Me.Cmd_annul.Enabled = False
   Me.Cmd_recfic.Enabled = False
   Me.Opt_imp.Enabled = False
   Me.Opt_word.Enabled = False
   Me.Opt_OOO.Enabled = False
   
    If Trim(Text2.Text) <> "" Then
        message = ""
'        If LCase(Left(Trim(text2.text), 3)) <> "c:\" Then
'            message = "Il faut renseigner le répertoire"
'        End If
        If message = "" And Mid(Right(Trim(Text2.Text), 4), 1, 1) <> "." Then
            Text2.Text = Text2.Text + ".doc"
        End If
        If message = "" And LCase(Right(Trim(Text2.Text), 4)) <> ".doc" Then
            message = "Il ne s'agit pas d'un fichier .doc ou l'extension n'est pas renseignée"
        End If
        If message <> "" Then
        reponse = MsgBox(message, , "Saisie du nom du fichier WORD ")
        Else
            nomword = Text2.Text
            Me.nom_fic = Text2.Text
            Set awd = Nothing
            Set ad = Nothing
            If trait_word Then
                MsgBox Label3.Caption, vbOKOnly, "Impression WORD"
            End If
            Unload Me
'                  FrmPrint.Show
            Unload FrmPrint
           
        End If
    Else
        reponse = MsgBox("Le nom n'est pas renseigné.", , "Saisie du nom du fichier WORD ")
    End If
ElseIf Opt_OOO Then
   Label3.Caption = ""
   Me.Cmd_ok.Enabled = False
   Me.Cmd_annul.Enabled = False
   Me.Cmd_recfic.Enabled = False
   Me.Opt_imp.Enabled = False
   Me.Opt_word.Enabled = False
   Me.Opt_OOO.Enabled = False
   
    If Trim(Text3.Text) <> "" Then
        message = ""
'        If LCase(Left(Trim(text2.text), 3)) <> "c:\" Then
'            message = "Il faut renseigner le répertoire"
'        End If
        If message = "" And Mid(Right(Trim(Text4.Text), 4), 1, 1) <> "." Then
            Text4.Text = Text4.Text + ".odt"
        End If
        If message = "" And LCase(Right(Trim(Text4.Text), 4)) <> ".odt" Then
            message = "Il ne s'agit pas d'un fichier .odt ou l'extension n'est pas renseignée"
        End If
        If message <> "" Then
        reponse = MsgBox(message, , "Saisie du nom du fichier OpenOffice ")
        Else
            nomword = Text3.Text + Text4.Text
            Me.nom_fic = Text3.Text + Text4.Text
            Set awd = Nothing
            Set ad = Nothing
            If trait_ooo Then
                MsgBox Label3.Caption, vbOKOnly, "Impression OpenOffice"
            End If
            Unload Me
'                  FrmPrint.Show
            Unload FrmPrint
           
        End If
    Else
        reponse = MsgBox("Le nom n'est pas renseigné.", , "Saisie du nom du fichier OpenOffice ")
    End If
End If
End Sub
Private Function trait_word() As Boolean
Dim ok_sing As Boolean
Dim npar As Integer, i As Integer
Dim arange As Word.Range
Dim f As Word.Frame
Dim titre2 As String, message As String
Dim nometude As String
nometude = "Etude : " + nom_etude
trait_word = True
If mod_save = "remplace" Then
   Kill (Me.Text2.Text)
End If
'Houpie 20040123 modif ajout messages
'            If Dir("c:\anohydro.txt") <> "" Then
'             '   exportTxt.Close
'                fso.DeleteFile ("c:\anohydro.txt")
'            End If
'            Set exportTxt = fso.CreateTextFile("c:\anohydro.txt", ForWriting)
'''''''''''''''''''''''''''fin modif
Me.MousePointer = 11
'  exportTxt.WriteLine "Début traitement (création fichier c:\anohydro.txt)"
On Error GoTo erreur
If awd Is Nothing Then
'    Set awd = New Word.Application
'  exportTxt.WriteLine "Avant création objet"
    Set awd = CreateObject("Word.Application")
'    Set awd = New Word.Application
'  exportTxt.WriteLine "Aprés création objet"
Else
'  exportTxt.WriteLine "Avant test  objet existe"
    If awd.Documents.count > 0 Then
        If Not ad Is Nothing Then
             ad.Close
        End If
    End If
    Set ad = Nothing
 ' exportTxt.WriteLine "Aprés test  objet existe"
End If

If Dir(nom_fic) <> "" Then
'  exportTxt.WriteLine "Avant ouverture document existant"
    Set ad = awd.Documents.Open(FileName:=nom_fic, ReadOnly:=False)
'  exportTxt.WriteLine "Aprés ouverture document existant  /  Avant cre_styles"
    Call cre_styles
'  exportTxt.WriteLine "Aprés cre_styles"
    npar = ad.Paragraphs.count - 1
'    message = "Avant ajout paragraphe " + Str(npar)
'  exportTxt.WriteLine message
    Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, _
        End:=ad.Paragraphs(npar).Range.End)
   If ad.Paragraphs.count > 1 Then
        Set myrange = ad.Range
        With myrange
            .Collapse Direction:=wdCollapseEnd
            .InsertBreak type:=wdSectionBreakNextPage  'wdPageBreak
        End With
    npar = npar + 1
    Else
        npar = 0
    End If
 ' exportTxt.WriteLine "Aprés ajout paragraphe"
Else
'  exportTxt.WriteLine "Avant nouveau document"
    Set ad = awd.Documents.Add
'  exportTxt.WriteLine "Aprés nouveau document  /  Avant cre_styles"
    Call cre_styles
'  exportTxt.WriteLine "Aprés cre_styles  /  Avant cadre_page"
    Call cadre_page
'  exportTxt.WriteLine "Aprés cadre_page  /  Avant cre_entete"
    Call cre_entete
    npar = 0
'  exportTxt.WriteLine "Aprés  cre_entete"
End If
'ad.Shapes.SelectAll
'Debug.Print ad.Shapes.Count
'ad.shapes.AddShape msoShapeRectangle, 5, 5, 10, 20
'ad.shapes.AddShape 96, 5, 5, 10, 20


'With ad.Shapes.AddShape(msoShapeRectangle, 0, 0, 5, 4).Fill
'    .ForeColor.RGB = RGB(128, 0, 0)
'    .BackColor.RGB = RGB(170, 170, 170)
'    .TwoColorGradient msoGradientHorizontal, 1
'End With
'  exportTxt.WriteLine "Avant sélection ouvrage"

titre2 = Me.nomobjet
If Type1 = "deversoir" Or Type1 = "deversoiror" Or Type1 = "pompe" Then
    titre2 = Me.nomobjet + "  --page 1/1--"
End If
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Range.InsertAfter nometude
npar = cre_titre(npar, Me.titre1, titre2)
Select Case Type1
    Case Is = "decant"
'      exportTxt.WriteLine "Sélection decant  / avant list_don1"
        npar = cre_table(npar, sstitre1, "list_don1", 4, 1)
'      exportTxt.WriteLine "aprés list_don1  / avant list_int1"
        npar = cre_table(npar, ssTitre2, "list_int1", 0, 1)
'      exportTxt.WriteLine "aprés list_int1  / avant list_resu1"
        npar = cre_table(npar, ssTitre3, "list_resu1", 0, 4)
'      exportTxt.WriteLine "aprés list_resu1  / avant dess.bmp"
        npar = recup_dess(npar, 449, 180, "dess.bmp")
'      exportTxt.WriteLine "aprés dess.bmp  / aprés Sélection decant"
     Case Is = "stockage"
        npar = cre_table(npar, sstitre1, "list_don1", 3, 0)
        npar = cre_table(npar, ssTitre2, "list_int1", 0, 0)
        npar = cre_table(npar, ssTitre3, "list_resu1", 0, 2)
        npar = recup_dess(npar, 447, 180, "dess.bmp")
     Case Is = "retention"
        npar = cre_table(npar, sstitre1, "list_don1", 1, 1)
        npar = cre_table(npar, ssTitre2, "list_int1", 0, 1)
        npar = recup_dess(npar, 449, 150, "dess1.bmp")
        npar = npar + 1
        ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
        ad.Range.InsertAfter Me.des2_titrb
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
'        ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
        npar = cre_table(npar, ssTitre3, "list_resu1", 0, 1)
        npar = recup_dess(npar, 449, 150, "dess.bmp")
    Case Is = "chute"
'''cre_tabled(npar,titre,nom du fichier,nb lignes vides avant,nb lignes vides aprés, _
    position texte champ3)
        npar = cre_tabled(npar, sstitre1, "list_don1", 4, 1, "Left")
        npar = cre_table(npar, ssTitre2, "list_don2", 0, 1)
        npar = cre_tabled(npar, ssTitre3, "list_int1", 0, 1, "Left")
        npar = cre_table(npar, ssTitre4, "list_resu1", 0, 4)
        npar = recup_dess(npar, 449, 180, "dess.bmp")
    Case Is = "conduite"
'''cre_tabled(npar,titre,nom du fichier,nb lignes vides avant,nb lignes vides aprés, _
    position texte champ3)
'   npar = cre_tabled(npar, sstitre1, "list_don1", 4, 1, "Left")
        npar = cre_table(npar, sstitre1, "list_don1", 6, 1)
        If Trim(ssTitre3) <> "" Then
            npar = cre_table(npar, ssTitre2, "list_don2", 0, 1)
            npar = cre_table(npar, ssTitre3, "list_int1", 0, 4)
        End If
        npar = recup_dess(npar, 449, 180, "dess.bmp")
     Case Is = "siphon"
        ok_sing = False
        npar = cre_tabled(npar, sstitre1, "list_don1", 4, 1, "Left")
        npar = cre_table(npar, ssTitre2, "list_don2", 0, 1)
        If Trim(ssTitre3) <> "" Then
            ok_sing = True
            npar = cre_tabled(npar, ssTitre3, "list_don3", 0, 1, "Right")
        End If
        If ok_sing Then
            npar = cre_table(npar, ssTitre4, "list_int1", 0, 2)
        Else
            npar = cre_table(npar, ssTitre4, "list_int1", 0, 4)
        End If
        npar = recup_dess(npar, 449, 180, "dess.bmp")
    Case Is = "versant"
        npar = cre_tablet(npar, sstitre1, "list_don1", 4, 1)
        npar = cre_tablet(npar, ssTitre2, "list_don2", 0, 1)
        npar = cre_tablet(npar, ssTitre3, "list_resu1", 0, 4)
        ad.Range.InsertAfter Me.des1_titrh
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
        ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
        npar = recup_dess(npar, 449, 180, "dess.bmp")
        npar = npar + 1
        ad.Range.InsertAfter Me.des1_titrb
        ad.Range.InsertParagraphAfter
        ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
        ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
        ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
    Case Is = "deversoir"
        npar = cre_table(npar, sstitre1, "list_don1", 3, 1) '4,1
        npar = cre_table(npar, ssTitre2, "list_don2", 1, 1)
        npar = cre_tableq(npar, ssTitre3, "list_don3", 1, 1)
        npar = cre_tableq(npar, ssTitre4, "list_don4", 1, 1) '3,1
        Set myrange = ad.Range
        With myrange
            .Collapse Direction:=wdCollapseEnd
            .InsertBreak type:=wdSectionBreakNextPage  'wdPageBreak
        End With
        npar = npar + 1
        titre2 = Me.nomobjet + "  --page 2/2--"
        npar = cre_titre(npar, Me.titre1, titre2)
        npar = cre_tabled(npar, ssTitre5, "list_don5", 3, 1, "Left") '4,1
        npar = recup_dess(npar, 449, 180, "dess.bmp")
        npar = npar + 1
        npar = cre_tabled(npar, ssTitre6, "list_don6", 1, 1, "Left") '2,1
        npar = recup_dess(npar, 449, 180, "dess1.bmp")
     Case Is = "deversoiror"
        npar = cre_table(npar, sstitre1, "list_don1", 3, 1) '4,1
'        npar = cre_table(npar, ssTitre2, "list_don2", 1, 1)
        npar = cre_tableq(npar, ssTitre3, "list_don3", 1, 1)
        npar = cre_table(npar, ssTitre4, "list_don4", 1, 1) '3,1
        npar = cre_table(npar, ssTitre5, "list_don5", 1, 1) '3,1
        Set myrange = ad.Range
        With myrange
            .Collapse Direction:=wdCollapseEnd
            .InsertBreak type:=wdSectionBreakNextPage  'wdPageBreak
        End With
        npar = npar + 1
        titre2 = Me.nomobjet + "  --page 2/2--"
        npar = cre_titre(npar, Me.titre1, titre2)
        npar = cre_table(npar, ssTitre6, "list_don6", 4, 4) '4,1
        npar = recup_dess(npar, 449, 300, "dess.bmp")
   Case Is = "pompe"
'        npar = cre_table(npar, sstitre1, "list_don1", 2, 0) '4,1
        npar = cre_tabled(npar, sstitre1, "list_don1", 3, 1, "Left") '4,1
        npar = cre_table(npar, ssTitre2, "list_don2", 1, 1)
 '       npar = cre_table(npar, ssTitre3, "list_don3", 1, 0)
        npar = cre_tableq(npar, ssTitre3, "list_don3", 1, 1)
        npar = cre_table(npar, ssTitre4, "list_don4", 1, 1)
        Set myrange = ad.Range
        With myrange
            .Collapse Direction:=wdCollapseEnd
            .InsertBreak type:=wdSectionBreakNextPage  'wdPageBreak
        End With
        npar = npar + 1
        titre2 = Me.nomobjet + "  --page 2/2--"
        npar = cre_titre(npar, Me.titre1, titre2)
        npar = cre_tabled(npar, ssTitre5, "list_don5", 2, 1, "Left")
'        npar = recup_dess(npar, 449, 180, "dess.bmp")
        npar = recup_dess(npar, 449, 350, "dess.bmp")
End Select
'certu 20080901 reporté en pied de page
'ad.Range.InsertParagraphAfter
'npar = npar + 1
'ad.Paragraphs(npar).Range.Style = ad.Styles(stylew3)
'ad.Range.InsertAfter "Dossier : " + nom_fich_edit

'  exportTxt.WriteLine "Aprés sélection ouvrage   / avant View"
'***********************************************
awd.Application.ActiveWindow.View.type = wdPageView
awd.Application.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
awd.Application.ActiveWindow.View.type = wdPageView
awd.Application.ActiveWindow.View.Zoom.Percentage = 75
'DoEvents
'awd.Application.Visible = True
DoEvents
'  exportTxt.WriteLine "Aprés  View"
 
ad.SaveAs nom_fic, wdFormatDocument
'  exportTxt.WriteLine "Avant message"
Select Case Trim(mod_save)
    Case Is = "remplace"
    Label3.Caption = " le fichier " + Text2.Text + " a été remplacé."
    Case Is = "complete"
    Label3.Caption = " le fichier " + Text2.Text + " a été complété."
    Case Is = ""
    Label3.Caption = " le fichier " + Text2.Text + " a été créé."
End Select
'  exportTxt.WriteLine "Aprés message"
ad.Close
Me.MousePointer = 1
' Houpie 20040123 modif ajout messages
'            exportTxt.Close
'            TotoTxt = "c:\anohydro.txt"
'            MyAppli = Shell("c:\windows\notepad.exe " & TotoTxt, vbNormalFocus)
'            AppActivate MyAppli
'''''''''''''fin modif
Exit Function
erreur:
    Me.MousePointer = 1
    MsgBox "Anomalie dans la création d'un fichier WORD", vbExclamation, "Impression WORD"
' Houpie 20040123 modif ajout messages
'            exportTxt.Close
'            TotoTxt = "c:\anohydro.txt"
'            MyAppli = Shell("c:\windows\notepad.exe " & TotoTxt, vbNormalFocus)
'            AppActivate MyAppli
'''''''''''''fin modif
trait_word = False
End Function
Private Function trait_ooo() As Boolean
Dim nomfich As String, nomfich0 As String
Dim Chemin As String, nomdir As String, nomdess As String, nomfich_dess As String
Dim liste_style() As String
Dim liste() As Variant
Dim xdate As String, xtime As String
Dim xdat1 As String, titre2 As String
Dim noPage As Integer
Dim sep As String, nomdir_pict As String, nomdir_meta As String
Dim aNomfich() As String, curdrive As String
Dim nomFichOdt As String
Dim chemin_pict As String, chemin_meta As String
Dim rep_pict As String, rep_meta As String, rep_etude As String
Dim Image1 As String, Image2 As String
nomFichOdt = Me.nom_fic
sep = Chr(34)
    trait_ooo = True
Me.MousePointer = 11
Image1 = ""
Image2 = ""
nomdir_pict = "Pictures"
nomdir_meta = "META-INF"
rep_pict = nomdir_pict + "\"
rep_meta = nomdir_meta + "\"
'  exportTxt.WriteLine "Début traitement (création fichier c:\anohydro.txt)"
On Error GoTo erreur
noPage = 1
titre2 = Me.nomobjet
If Type1 = "deversoir" Or Type1 = "deversoiror" Or Type1 = "pompe" Then
    titre2 = Me.nomobjet + "  --page 1/1--"
End If
xdate = Date
xtime = Time
'    nomdir = Left$(Me.Text2.Text, Len(Trim$(Me.Text2.Text)) - 4)
'    Chemin = nomdir + "\"
    Chemin = chemin_etude
    rep_etude = Left$(Chemin, Len(Chemin) - 1)
'    curdrive = Left$(Chemin, 3)
        chemin_pict = Chemin + nomdir_pict
        If Dir(chemin_pict, vbDirectory) = "" Then
            MkDir (chemin_pict)
        Else
            Call del_fich_rep(chemin_pict)
        End If
        chemin_meta = Chemin + nomdir_meta
        If Dir(chemin_meta, vbDirectory) = "" Then
            MkDir (chemin_meta)
        Else
            Call del_fich_rep(chemin_meta)
        End If
    nomfich = "styles.xml"
    ReDim Preserve aNomfich(1)
    aNomfich(UBound(aNomfich)) = nomfich
    nomfich = Chemin + "styles.xml"
    Call Ecr_styles(nomfich, xdate)
    nomfich = "meta.xml"
    ReDim Preserve aNomfich(UBound(aNomfich) + 1)
    aNomfich(UBound(aNomfich)) = nomfich
    xdat1 = Mid(xdate, 7, 4) + "-" + Mid(xdate, 4, 2) + "-" + Mid(xdate, 1, 2) + "T" + Mid(xtime, 1, 2) + ":" + Mid(xtime, 4, 2) + ":" + Mid(xtime, 7, 2)
    nomfich = Chemin + "meta.xml"
    Call Ecr_meta(nomfich, xdat1)
    nomfich = "content.xml"
    ReDim Preserve aNomfich(UBound(aNomfich) + 1)
    aNomfich(UBound(aNomfich)) = nomfich
    xdat1 = Mid(xdate, 1, 2) + Mid(xdate, 4, 2) + Mid(xdate, 7, 4) + "_" + Mid(xtime, 1, 2) + Mid(xtime, 4, 2) + Mid(xtime, 7, 2)
    nomfich = Chemin + "content.xml"
    If mod_save = "remplace" Then
       Kill (nomFichOdt)   'fichier zip (odt)
'        If exist_fich(nomdir) Then
'            Call del_fich_rep(nomdir)
'        End If
'        If Dir(nomdir, vbDirectory) <> "" Then
'            RmDir (nomdir)
'        End If
        FileCopy chemin_app + "ini_odt.gab", nomFichOdt
'        MkDir (nomdir)

 '       If Dir(nomfich) <> "" Then
 '           Kill nomfich      'fichier content.xml
 '       End If
    ElseIf mod_save = "" Then
        FileCopy chemin_app + "ini_odt.gab", nomFichOdt
'        MkDir (nomdir)
'    ElseIf mod_save = "complete" Then
'        FileCopy nomfich, nomfich0
'        Kill nomfich      'fichier content.xml
   End If

    lhFicooo1 = FreeFile
    Open nomfich For Output As #lhFicooo1
'    If mod_save = "complete" Then
 '       Call recopy_debut(nomfich0)
 '       Kill nomfich0
 '   Else
        Call Ecr_debut
 '   End If
    Index = 10
    Call Ecr_titre(titre1, titre2, noPage)
    Select Case Type1
       Case Is = "decant"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(4, sstitre1)  ' 4 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_int1")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre3)
            liste = owner.fobjet.lect_list("list_resu1")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "7.00", nomdess)
            Image1 = nomdess
        Case Is = "stockage"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(3, sstitre1)  ' 3 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_int1")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre3)
            liste = owner.fobjet.lect_list("list_resu1")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "7.00", nomdess)
            Image1 = nomdess
        Case Is = "retention"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, sstitre1)  ' 1 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_int1")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess1" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess1.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess1" + xdat1 + ".bmp"
            Call Ecr_dess(0, "4.50", nomdess)
            Image1 = nomdess
            chaine = "<text:p text:style-name=" & sep & "P1" & sep & ">"
            Print #lhFicooo1, chaine
            chaine = "<text:span text:style-name=" & sep & "T3" & sep & ">" & Me.des2_titrb & "</text:span>"
            Print #lhFicooo1, chaine
            chaine = "</text:p>"
            Print #lhFicooo1, chaine
            Call Ecr_titre_tableau(0, ssTitre3)
            liste = owner.fobjet.lect_list("list_resu1")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(0, "4.50", nomdess)
            Image2 = nomdess
         Case Is = "chute"
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(3, sstitre1)  ' 3 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_don2")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre3)  ' 1 = nb lignes avant
            liste = owner.fobjet.lect_list("list_int1")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre4)
            liste = owner.fobjet.lect_list("list_resu1")
            Call Ecr_tableau("4" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "7.00", nomdess)
            Image1 = nomdess
         Case Is = "conduite"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(4, sstitre1)  ' 4 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            If Trim(ssTitre3) <> "" Then
                Call Ecr_titre_tableau(1, ssTitre2)
                liste = owner.fobjet.lect_list("list_don2")
                Call Ecr_tableau("2" + xdat1, liste, liste_style)
                Call Ecr_titre_tableau(1, ssTitre3)
                liste = owner.fobjet.lect_list("list_int1")
                Call Ecr_tableau("3" + xdat1, liste, liste_style)
            End If
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "7.00", nomdess)
            Image1 = nomdess
        Case Is = "siphon"
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(2, sstitre1)  ' 2 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_don2")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            If Trim(ssTitre3) <> "" Then
                ReDim liste_style(5, 2)
                liste_style(1, 1) = "A"
                liste_style(2, 1) = "B"
                liste_style(3, 1) = "C"
                liste_style(4, 1) = "D"
                liste_style(5, 1) = "E"
                liste_style(1, 2) = "P6"
                liste_style(2, 2) = "P5"
                liste_style(3, 2) = "P5"
                liste_style(4, 2) = "P5"
                liste_style(5, 2) = "P4"
                Call Ecr_titre_tableau(1, ssTitre3)  ' 1 = nb lignes avant
                liste = owner.fobjet.lect_list("list_don3")
                Call Ecr_tableau("3" + xdat1, liste, liste_style)
            End If
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre4)
            liste = owner.fobjet.lect_list("list_int1")
            Call Ecr_tableau("4" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "7.00", nomdess)
            Image1 = nomdess
        Case Is = "versant"
            ReDim liste_style(6, 2)
            liste_style(1, 1) = "I"
            liste_style(2, 1) = "E"
            liste_style(3, 1) = "E"
            liste_style(4, 1) = "I"
            liste_style(5, 1) = "E"
            liste_style(6, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P4"
            liste_style(5, 2) = "P5"
            liste_style(6, 2) = "P4"
            Call Ecr_titre_tableau(2, sstitre1)  ' 2 = nb lignes avant
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_don2")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre3)
            liste = owner.fobjet.lect_list("list_resu1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            chaine = "<text:p text:style-name=" & sep & "Standard" & sep & "/>"
            Print #lhFicooo1, chaine
            chaine = "<text:p text:style-name=" & sep & "P1" & sep & ">"
            Print #lhFicooo1, chaine
            chaine = "<text:span text:style-name=" & sep & "T3" & sep & ">" & Me.des1_titrh & "</text:span>"
'            chaine = "<text:p text:style-name=" & sep & "P3" & sep & ">" & Me.des1_titrh & "</text:p>"
            Print #lhFicooo1, chaine
            chaine = "</text:p>"
            Print #lhFicooo1, chaine
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(0, "7.00", nomdess)
            Image1 = nomdess
            chaine = "<text:p text:style-name=" & sep & "Standard" & sep & "/>"
            Print #lhFicooo1, chaine
            chaine = "<text:p text:style-name=" & sep & "P1" & sep & ">"
            Print #lhFicooo1, chaine
            chaine = "<text:span text:style-name=" & sep & "T3" & sep & ">" & Me.des1_titrb & "</text:span>"
'            chaine = "<text:p text:style-name=" & sep & "P3" & sep & ">" & Me.des1_titrb & "</text:p>"
            Print #lhFicooo1, chaine
            chaine = "</text:p>"
            Print #lhFicooo1, chaine
        Case Is = "deversoir"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(2, sstitre1)
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_don2")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            ReDim liste_style(7, 2)
            liste_style(1, 1) = "G"
            liste_style(2, 1) = "E"
            liste_style(3, 1) = "E"
            liste_style(4, 1) = "E"
            liste_style(5, 1) = "E"
            liste_style(6, 1) = "H"
            liste_style(7, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            liste_style(6, 2) = "P5"
            liste_style(7, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre3)
            liste = owner.fobjet.lect_list("list_don3")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(1, ssTitre4)
            liste = owner.fobjet.lect_list("list_don4")
            Call Ecr_tableau("4" + xdat1, liste, liste_style)
            chaine = "<text:p text:style-name=" & sep & "P9" & sep & " />"
            Print #lhFicooo1, chaine
            titre2 = Me.nomobjet + "  --page 2/2--"
            noPage = noPage + 1
            Call Ecr_titre(titre1, titre2, noPage)
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(2, ssTitre5)
            liste = owner.fobjet.lect_list("list_don5")
            Call Ecr_tableau("5" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "4.50", nomdess)
            Image1 = nomdess
            Call Ecr_titre_tableau(1, ssTitre6)
            liste = owner.fobjet.lect_list("list_don6")
            Call Ecr_tableau("6" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess1" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess1.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess1" + xdat1 + ".bmp"
            Call Ecr_dess(1, "4.50", nomdess)
            Image2 = nomdess
       Case Is = "deversoiror"
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(2, sstitre1)
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            ReDim liste_style(7, 2)
            liste_style(1, 1) = "G"
            liste_style(2, 1) = "E"
            liste_style(3, 1) = "E"
            liste_style(4, 1) = "E"
            liste_style(5, 1) = "E"
            liste_style(6, 1) = "H"
            liste_style(7, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            liste_style(6, 2) = "P5"
            liste_style(7, 2) = "P4"
            Call Ecr_titre_tableau(2, ssTitre3)
            liste = owner.fobjet.lect_list("list_don3")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(2, ssTitre4)
            liste = owner.fobjet.lect_list("list_don4")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            Call Ecr_titre_tableau(2, ssTitre5)
            liste = owner.fobjet.lect_list("list_don5")
            Call Ecr_tableau("4" + xdat1, liste, liste_style)
            chaine = "<text:p text:style-name=" & sep & "P9" & sep & " />"
            Print #lhFicooo1, chaine
            titre2 = Me.nomobjet + "  --page 2/2--"
            noPage = noPage + 1
            Call Ecr_titre(titre1, titre2, noPage)
            Call Ecr_titre_tableau(4, ssTitre6)
            liste = owner.fobjet.lect_list("list_don6")
            Call Ecr_tableau("5" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(4, "10.00", nomdess)
            Image1 = nomdess
        Case Is = "pompe"
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(2, sstitre1)
            liste = owner.fobjet.lect_list("list_don1")
            Call Ecr_tableau("1" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre2)
            liste = owner.fobjet.lect_list("list_don2")
            Call Ecr_tableau("2" + xdat1, liste, liste_style)
            ReDim liste_style(6, 2)
            liste_style(1, 1) = "J"
            liste_style(2, 1) = "E"
            liste_style(3, 1) = "E"
            liste_style(4, 1) = "E"
            liste_style(5, 1) = "E"
            liste_style(6, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P10"
            liste_style(3, 2) = "P10"
            liste_style(4, 2) = "P10"
            liste_style(5, 2) = "P10"
            liste_style(6, 2) = "P10"
            Call Ecr_titre_tableau(1, ssTitre3)
            liste = owner.fobjet.lect_list("list_don3")
            Call Ecr_tableau("3" + xdat1, liste, liste_style)
            ReDim liste_style(3, 2)
            liste_style(1, 1) = "F"
            liste_style(2, 1) = "D"
            liste_style(3, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            Call Ecr_titre_tableau(1, ssTitre4)
            liste = owner.fobjet.lect_list("list_don4")
            Call Ecr_tableau("4" + xdat1, liste, liste_style)
            chaine = "<text:p text:style-name=" & sep & "P9" & sep & " />"
            Print #lhFicooo1, chaine
            titre2 = Me.nomobjet + "  --page 2/2--"
            noPage = noPage + 1
            Call Ecr_titre(titre1, titre2, noPage)
            ReDim liste_style(5, 2)
            liste_style(1, 1) = "A"
            liste_style(2, 1) = "B"
            liste_style(3, 1) = "C"
            liste_style(4, 1) = "D"
            liste_style(5, 1) = "E"
            liste_style(1, 2) = "P6"
            liste_style(2, 2) = "P5"
            liste_style(3, 2) = "P4"
            liste_style(4, 2) = "P5"
            liste_style(5, 2) = "P4"
            Call Ecr_titre_tableau(2, ssTitre5)
            liste = owner.fobjet.lect_list("list_don5")
            Call Ecr_tableau("5" + xdat1, liste, liste_style)
            nomdess = rep_pict + "dess" + xdat1 + ".bmp"
            nomfich_dess = Chemin + nomdess
            FileCopy chemin_app + "dess.bmp", nomfich_dess
            ReDim Preserve aNomfich(UBound(aNomfich) + 1)
'           aNomfich(UBound(aNomfich)) = nomfich_dess
            aNomfich(UBound(aNomfich)) = nomdess
            nomdess = nomdir_pict + "/dess" + xdat1 + ".bmp"
            Call Ecr_dess(1, "10.00", nomdess)
            Image1 = nomdess
End Select

    Call Ecr_fin
    
    Close #lhFicooo1
    nomfich = rep_meta + "manifest.xml"
    ReDim Preserve aNomfich(UBound(aNomfich) + 1)
    aNomfich(UBound(aNomfich)) = nomfich
    nomfich = Chemin + rep_meta + "manifest.xml"
    Call Ecr_manifest(nomfich, Image1, Image2)

'****************************
' ajouter les fichiers de nomdir dans le fichier odt : nom
   Dim ret As Boolean
   Dim spath As String
   spath = CurDir()
   ChDir (rep_etude)
   ret = AddODT(nomFichOdt, aNomfich)
'******************************
  ' nom = nomfich
   
    Call del_fichiers(aNomfich)
    ChDir (spath)
    Call del_fich_rep(chemin_pict)
    Call del_fich_rep(chemin_meta)
    If Dir(chemin_pict, vbDirectory) <> "" Then
        RmDir (chemin_pict)
    End If
    If Dir(chemin_meta, vbDirectory) <> "" Then
        RmDir (chemin_meta)
    End If

   
    Select Case Trim(mod_save)
    Case Is = "remplace"
    Label3.Caption = " le fichier " + nomFichOdt + " a été remplacé."
    Case Is = "complete"
    Label3.Caption = " le fichier " + nomFichOdt + " a été complété."
    Case Is = ""
    Label3.Caption = " le fichier " + nomFichOdt + " a été créé."
    End Select
    Me.MousePointer = 1
Exit Function
erreur:
    Me.MousePointer = 1
    MsgBox "Anomalie dans la création d'un fichier OpenOffice", vbExclamation, "Impression OpenOffice"
trait_ooo = False

End Function

Function style_existe(ad As Word.Document, nstyle As String) As Boolean
Dim i As Integer, ns As Integer, lastyl As Word.Styles, astyl As Word.Style

On Error GoTo trait_erreur
style_existe = False
'Set lastyl = ad.Styles
'ns = lastyl.Count
ns = ad.Styles.count
For i = 1 To ns
'Set astyl = lastyl(i)
'If astyl.NameLocal = nstyle Then
If ad.Styles(i).NameLocal = nstyle Then
    style_existe = True
    Exit For
End If
Next
Exit Function
trait_erreur:
    MsgBox "Anomalie dans le test d'existence des styles", vbExclamation, "Impression WORD"
End Function
Private Sub cre_styles()
'    wrstyles = awd.Languages(wdFrench).WritingStyleList 'wdFrench =1036
On Error GoTo trait_erreur
    Set stylew1 = ad.Styles("normal")
    stylew1.Font.Size = 10
    stylew1.Font.Bold = False
    stylew1.Font.Italic = False
If Not style_existe(ad, "h_sstitre") Then
     Set stylew0 = ad.Styles("normal")
      ad.Styles.Add "h_sstitre", stylew0.type
    Set stylew0 = ad.Styles("h_sstitre")
        stylew0.Font.Size = 11
        stylew0.Font.Bold = True
        stylew0.Font.Italic = False
Else
    Set stylew0 = ad.Styles("h_sstitre")
End If
If Not style_existe(ad, "h_paragraphe") Then
    Set stylew2 = ad.Styles("normal")
    ad.Styles.Add "h_paragraphe", stylew2.type
    Set stylew2 = ad.Styles("h_paragraphe")
    stylew2.Font.Size = 11
    stylew2.Font.Bold = True
    stylew2.Font.Italic = False
Else
    Set stylew2 = ad.Styles("h_paragraphe")
End If
If Not style_existe(ad, "h_titre") Then
    Set stylew3 = ad.Styles("normal")
    ad.Styles.Add "h_titre", stylew3.type
    Set stylew3 = ad.Styles("h_titre")
    stylew3.Font.Size = 12
    stylew3.Font.Bold = True
    stylew3.Font.Italic = False
Else
    Set stylew3 = ad.Styles("h_titre")
End If
If Not style_existe(ad, "h_entete") Then
    Set stylew4 = ad.Styles("normal") '(73)
    ad.Styles.Add "h_entete", stylew4.type
    Set stylew4 = ad.Styles("h_entete")
    stylew4.Font.Size = 12
    stylew4.Font.Bold = True
    stylew4.Font.Italic = False
Else
    Set stylew4 = ad.Styles("h_entete")
End If
If Not style_existe(ad, "h_titregraphique") Then
    Set stylew = ad.Styles("normal") '49
    ad.Styles.Add "h_titregraphique", stylew.type
    Set stylew = ad.Styles("h_titregraphique") '49
    stylew.Font.Size = 10
    stylew.Font.Bold = False
    stylew.Font.Italic = False
Else
    Set stylew = ad.Styles("h_titregraphique") '49
End If

Exit Sub
trait_erreur:
    MsgBox "Anomalie dans la définition des styles", vbExclamation, "Impression WORD"
End Sub
Private Sub cadre_page()
On Error GoTo trait_erreur
   With ad.Sections(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .ColorIndex = wdAuto
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
            .ColorIndex = wdAuto
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
'    With Options
'        .DefaultBorderLineStyle = wdLineStyleSingle
'        .DefaultBorderLineWidth = wdLineWidth150pt
'        .DefaultBorderColorIndex = wdAuto
'    End With
Exit Sub
trait_erreur:
    MsgBox "Anomalie dans la création du cadre d'une page", vbExclamation, "Impression WORD"

End Sub

Private Function cre_titre(ByVal npar As Integer, ByVal titre1 As String, _
    ByVal titre2 As String) As Integer
Dim arange As Word.Range
Dim f As Word.Frame
Dim npard As Integer, nparf As Integer
Dim nometude As String
'    nometude = "Etude : " + nom_etude
On Error GoTo trait_erreur
ad.Range.InsertParagraphAfter
npar = npar + 1
npard = npar
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
ad.Range.InsertAfter titre1
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew3)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
ad.Range.InsertAfter titre2
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew0)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphCenter
nparf = npar
Set arange = ad.Range(Start:=ad.Paragraphs(npard).Range.End, _
    End:=ad.Paragraphs(nparf).Range.End)
With arange.Shading
    .Texture = wdTexture20Percent 'fond en gris
End With
'Set f = ad.Frames.Add(arange) 'ajouter un cadre autour d'un paragraphe
'With f.Shading
'    .Texture = wdTexture20Percent 'fond en gris
'End With
'    f.Borders.Enable = False
'Set arange = ad.Range(Start:=ad.Paragraphs(npard).Range.Start, _
'    End:=ad.Paragraphs(nparf).Range.End)
'Set f = ad.Range.Frames.Add(arange) 'ajouter un cadre autour d'un paragraphe
'    f.Borders.Enable = False
cre_titre = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la création du titre", vbExclamation, "Impression WORD"
End Function
Private Sub cre_entete()
Dim s1 As Double, s2 As Double
Dim hd As Word.HeaderFooter
Dim mrange As Word.Range
Dim np As Integer
Dim Sh As Word.Shape
On Error GoTo trait_erreur
np = 1
    ad.Sections(1).Footers(wdHeaderFooterPrimary).Range _
     .InsertDateTime , _
     InsertAsField:=True
'     .InsertDateTime DateTimeFormat:="jj MMMM aaaa", _
'     InsertAsField:=True
    With ad.Sections(1).Footers(wdHeaderFooterPrimary)
        .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight
        .Range.InsertAfter " - Dossier : " + nom_fich_edit
    End With
    Set hd = ad.Sections(1).Headers(wdHeaderFooterPrimary)
    With hd
        Set mrange = hd.Range
        s1 = mrange.Start
       Set Sh = .Shapes.AddPicture(FileName:=chemin_app + "texte.bmp", _
         LinkToFile:=False, SaveWithDocument:=True)
         Sh.Left = 0
         Sh.Top = 5
         Sh.Height = 55
        With mrange
            .Start = s1
 '           s1 = .Start
            .InsertAfter _
                  "Boite à Outils Hydrologie , Hydraulique et Assainissement"
            .Paragraphs(np).Style = ad.Styles(stylew4) '(stylew4)
 
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .InsertParagraphAfter
            np = np + 1
            .InsertParagraphAfter
            np = np + 1
            .Paragraphs(np).Style = ad.Styles(stylew1) '(stylew4)
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .ParagraphFormat.FirstLineIndent = 100
            .ParagraphFormat.LeftIndent = 0
            .InsertAfter text_serv1 '"Centre d'études Techniques de l' Equipement de l' Est"
            .InsertParagraphAfter
            np = np + 1
            .InsertAfter text_serv2 '"Laboratoire Régional de Nancy"
            .InsertParagraphAfter
   '          np = np + 1
'            .Paragraphs(np).Style = ad.Styles(stylew1) '(stylew4)
'            .ParagraphFormat.Alignment = wdAlignParagraphLeft
'            .ParagraphFormat.FirstLineIndent = 100
'            .ParagraphFormat.LeftIndent = 0
       End With
        s2 = mrange.End
    End With
    With mrange
'    .Start = s1
'    .End = s2
    With .Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt 'Options.DefaultBorderLineWidth
        .ColorIndex = wdBlack
    End With
    With .Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt 'Options.DefaultBorderLineWidth
        .ColorIndex = wdBlack
    End With
    With .Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt 'Options.DefaultBorderLineWidth
        .ColorIndex = wdBlack
    End With
    With .Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle 'Options.DefaultBorderLineStyle
        .LineWidth = wdLineWidth150pt 'Options.DefaultBorderLineWidth
        .ColorIndex = wdBlack
    End With
    End With
Exit Sub
trait_erreur:
    MsgBox "Anomalie dans la création de l'en-tête", vbExclamation, "Impression WORD"

End Sub
Private Function cre_table(ByVal npar As Integer, ByVal Titre As String, _
    ByVal nomfich As String, ByVal nlav As Integer, ByVal nlap As Integer) As Integer
Dim liste() As Variant
Dim nb As Integer, i As Integer
Dim ct As Word.Table
Dim xlar As Double
Dim mycell As Selection
On Error GoTo trait_erreur
liste = owner.fobjet.lect_list(nomfich)
If nlav > 0 Then
    For i = 1 To nlav
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
nb = UBound(liste)
ad.Range.InsertAfter Titre
ad.Range.InsertParagraphAfter
npar = npar + 1
npar = ad.Paragraphs.count - 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew2)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, End:=ad.Paragraphs(npar).Range.End)
'************création du tableau
Set ct = ad.Tables.Add(Range:=arange, NumRows:=nb + 1, NumColumns:=4)
ct.Columns(1).SetWidth ColumnWidth:=42.5, RulerStyle:=wdAdjustProportional
ct.Columns(2).SetWidth ColumnWidth:=255.1, RulerStyle:=wdAdjustProportional
ct.Columns(3).SetWidth ColumnWidth:=113.4, RulerStyle:=wdAdjustProportional
ct.Borders(wdBorderHorizontal) = False
ct.Borders(wdBorderVertical) = False
       For i = 1 To nb + 1
            ct.Cell(i, 2).Range.Text = liste(i - 1, 1)
            ct.Cell(i, 3).Range.Text = liste(i - 1, 2)
            ct.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            ct.Cell(i, 4).Range.Text = liste(i - 1, 3)
       Next
With ct
    With .Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
End With
    npar = ad.Paragraphs.count - 1
If nlap > 0 Then
    For i = 1 To nlap
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
    cre_table = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la création d'un tableau (4 colonnes : cre_table)", vbExclamation, "Impression WORD"
End Function
Private Function cre_tabled(ByVal npar As Integer, ByVal Titre As String, _
    ByVal nomfich As String, ByVal nlav As Integer, ByVal nlap As Integer, _
    ByVal pos As String) As Integer
Dim liste() As Variant
Dim nb As Integer, i As Integer
Dim ct As Word.Table
Dim xlar As Double
Dim mycell As Selection
On Error GoTo trait_erreur
liste = owner.fobjet.lect_list(nomfich)
If nlav > 0 Then
    For i = 1 To nlav
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
nb = UBound(liste)
ad.Range.InsertAfter Titre
ad.Range.InsertParagraphAfter
npar = npar + 1
npar = ad.Paragraphs.count - 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew2)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, End:=ad.Paragraphs(npar).Range.End)
'************création du tableau
Set ct = ad.Tables.Add(Range:=arange, NumRows:=nb + 1, NumColumns:=6)
ct.Columns(1).SetWidth ColumnWidth:=20#, RulerStyle:=wdAdjustProportional
ct.Columns(2).SetWidth ColumnWidth:=170#, RulerStyle:=wdAdjustProportional
ct.Columns(3).SetWidth ColumnWidth:=80#, RulerStyle:=wdAdjustProportional
ct.Columns(4).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(5).SetWidth ColumnWidth:=80#, RulerStyle:=wdAdjustProportional
ct.Borders(wdBorderHorizontal) = False
ct.Borders(wdBorderVertical) = False
       For i = 1 To nb + 1
            ct.Cell(i, 2).Range.Text = liste(i - 1, 1)
            ct.Cell(i, 3).Range.Text = liste(i - 1, 2)
            ct.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            ct.Cell(i, 4).Range.Text = liste(i - 1, 3)
            If pos = "Right" Then
            ct.Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            End If
            ct.Cell(i, 5).Range.Text = liste(i - 1, 4)
            ct.Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            ct.Cell(i, 6).Range.Text = liste(i - 1, 5)
       Next
With ct
    With .Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
End With
    npar = ad.Paragraphs.count - 1
If nlap > 0 Then
    For i = 1 To nlap
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
    cre_tabled = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la création d'un tableau (6 colonnes : cre_tabled)", vbExclamation, "Impression WORD"
End Function
Private Function cre_tablet(ByVal npar As Integer, ByVal Titre As String, _
    ByVal nomfich As String, ByVal nlav As Integer, ByVal nlap As Integer) As Integer
Dim liste() As Variant
Dim nb As Integer, i As Integer
Dim ct As Word.Table
Dim xlar As Double
Dim mycell As Selection
'If Type1 = "versant" Then
On Error GoTo trait_erreur
If Not owner.fbassin Is Nothing Then
    liste = owner.fbassin.lect_list(nomfich)
Else
    liste = owner.fobjet.lect_list(nomfich)
End If
If nlav > 0 Then
    For i = 1 To nlav
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
nb = UBound(liste)
ad.Range.InsertAfter Titre
ad.Range.InsertParagraphAfter
npar = npar + 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew2)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, End:=ad.Paragraphs(npar).Range.End)
'************création du tableau
Set ct = ad.Tables.Add(Range:=arange, NumRows:=nb + 1, NumColumns:=7)
ct.Columns(1).SetWidth ColumnWidth:=10#, RulerStyle:=wdAdjustProportional
ct.Columns(2).SetWidth ColumnWidth:=120#, RulerStyle:=wdAdjustProportional
ct.Columns(3).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(4).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(5).SetWidth ColumnWidth:=120#, RulerStyle:=wdAdjustProportional
ct.Columns(6).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Borders(wdBorderHorizontal) = False
ct.Borders(wdBorderVertical) = False
       For i = 1 To nb + 1
            ct.Cell(i, 2).Range.Text = liste(i - 1, 1)
            ct.Cell(i, 3).Range.Text = liste(i - 1, 2)
            ct.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            ct.Cell(i, 4).Range.Text = liste(i - 1, 3)
            ct.Cell(i, 5).Range.Text = liste(i - 1, 4)
            ct.Cell(i, 6).Range.Text = liste(i - 1, 5)
            ct.Cell(i, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            ct.Cell(i, 7).Range.Text = liste(i - 1, 6)
       Next
With ct
    With .Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
End With
    npar = ad.Paragraphs.count - 1
If nlap > 0 Then
    For i = 1 To nlap
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
    cre_tablet = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la création d'un tableau (7 colonnes : cre_tablet)", vbExclamation, "Impression WORD"
End Function
Private Function cre_tableq(ByVal npar As Integer, ByVal Titre As String, _
    ByVal nomfich As String, ByVal nlav As Integer, ByVal nlap As Integer) As Integer
Dim liste() As Variant
Dim nb As Integer, i As Integer
Dim ct As Word.Table
Dim xlar As Double
Dim mycell As Selection
On Error GoTo trait_erreur
liste = owner.fobjet.lect_list(nomfich)
If nlav > 0 Then
    For i = 1 To nlav
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
nb = UBound(liste)
ad.Range.InsertAfter Titre
ad.Range.InsertParagraphAfter
npar = npar + 1
npar = ad.Paragraphs.count - 1
ad.Paragraphs(npar).Range.Style = ad.Styles(stylew2)
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceBefore = 0
ad.Paragraphs(npar).Range.ParagraphFormat.SpaceAfter = 0
ad.Paragraphs(npar).Alignment = wdAlignParagraphLeft
Set arange = ad.Range(Start:=ad.Paragraphs(npar).Range.End, End:=ad.Paragraphs(npar).Range.End)
'************création du tableau
Set ct = ad.Tables.Add(Range:=arange, NumRows:=nb + 1, NumColumns:=8)
ct.Columns(1).SetWidth ColumnWidth:=10#, RulerStyle:=wdAdjustProportional
ct.Columns(2).SetWidth ColumnWidth:=150#, RulerStyle:=wdAdjustProportional  '130
ct.Columns(3).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(4).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(5).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(6).SetWidth ColumnWidth:=50#, RulerStyle:=wdAdjustProportional
ct.Columns(7).SetWidth ColumnWidth:=60#, RulerStyle:=wdAdjustProportional
ct.Borders(wdBorderHorizontal) = False
ct.Borders(wdBorderVertical) = False
For i = 1 To nb + 1
     ct.Cell(i, 2).Range.Text = liste(i - 1, 1)
     ct.Cell(i, 3).Range.Text = liste(i - 1, 2)
     ct.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    If Type1 = "pompe" Then
         ct.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
     ct.Cell(i, 4).Range.Text = liste(i - 1, 3)
    If Type1 = "pompe" Then
         ct.Cell(i, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
     ct.Cell(i, 5).Range.Text = liste(i - 1, 4)
     ct.Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    If Type1 = "pompe" Then
         ct.Cell(i, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
     ct.Cell(i, 6).Range.Text = liste(i - 1, 5)
    If Type1 = "pompe" Then
         ct.Cell(i, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
     ct.Cell(i, 7).Range.Text = liste(i - 1, 6)
     ct.Cell(i, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
     If Type1 = "pompe" Then
         ct.Cell(i, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    ct.Cell(i, 8).Range.Text = liste(i - 1, 7)
    If Type1 = "pompe" Then
         ct.Cell(i, 8).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If

Next
With ct
    With .Borders(wdBorderLeft)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderRight)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
    With .Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .ColorIndex = wdAuto
    End With
End With
    npar = ad.Paragraphs.count - 1
If nlap > 0 Then
    For i = 1 To nlap
        ad.Range.InsertParagraphAfter
        npar = npar + 1
        ad.Paragraphs(npar).Range.Style = ad.Styles(stylew)
    Next
End If
    cre_tableq = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la création d'un tableau (8 colonnes : cre_tableq)", vbExclamation, "Impression WORD"
End Function
Private Function recup_dess(ByVal npar As Integer, ByVal xlar As String, ByVal haut As String, _
    ByVal nomfich As String) As Integer
    Dim f As Word.Frame
    Dim Sh As Word.InlineShape
'**cree un cadre pour insérer un dessin*****
    On Error GoTo trait_erreur
        Set maplage = ad.Range(Start:=ad.Paragraphs(npar).Range.End, _
        End:=ad.Paragraphs(npar).Range.End)
'        Set f = ad.Frames.Add(Range:=maplage)
'
'        f.Width = xlar
'        f.Height = haut
 '*** insertion du dessin*****
        Set Sh = maplage.InlineShapes.AddPicture(FileName:=chemin_app + nomfich, _
                 LinkToFile:=False, SaveWithDocument:=True)
'                 LinkToFile:=True, SaveWithDocument:=True)
        Sh.LockAspectRatio = False
        Sh.Height = haut
        Sh.Width = xlar
        With Sh
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle 'Options.DefaultBorderLineStyle
            .LineWidth = wdLineWidth150pt
    '        .LineWidth = Options.DefaultBorderLineWidth
    '        .ColorIndex = Options.DefaultBorderColorIndex
        End With
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        End With
'        ad.Range.InsertParagraphAfter
'        npar = npar + 1
    recup_dess = npar
Exit Function
trait_erreur:
    MsgBox "Anomalie dans la récupération d'un schéma (recup_dess)", vbExclamation, "Impression WORD"
End Function



Private Sub Cmd_repos_Click()
Dim reponse As Integer
Dim message As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim f As File
Dim d As Drive
Dim nom As String
sav_word = True
'    Set frmf = Nothing


fsco = create_fs(cdlg1.FileName)
's = fsco.nom
Dim chemin_imp As String
chemin_imp = Text3.Text 'fsco.lecteur + fsco.Chemin
s = Text4.Text + ".odt"
    While sav_word
        On Error GoTo errHandler
        cdlg1.DialogTitle = "Enregistrer sous " '"Recherche d'un fichier "
        cdlg1.FileName = s '""
        cdlg1.Filter = "Fichiers OpenOffice (*.odt)|*.odt|Tous (*.*)|*.*"
        cdlg1.InitDir = chemin_imp ' chemin_etude
        cdlg1.Flags = cdlOFNHideReadOnly
        cdlg1.CancelError = True
        
        'cdlg1.ShowOpen
        cdlg1.ShowSave
        s = cdlg1.FileName
        On Error GoTo 0
        If s <> "" Then
        fsco = create_fs(s)
        If fsco.dr_type = 1 Then
            message = "Fichier sur disquette;" + Chr(13) + Chr(10) + "Vérifier que la disquette n'est pas protégée en écriture."
            reponse = MsgBox(message, , "Saisie du nom du fichier OpenOffice ")
        End If
        If fsco.dr_type = 4 Then
            message = "Fichier sur CR-ROM;" + Chr(13) + Chr(10) + "Pas d'accés en écriture."
            reponse = MsgBox(message, , "Saisie du nom du fichier OpenOffice ")
            nom = ""
         '   Text1.Text = ""
        '        Text2.Text = ""
        ElseIf fsco.lecteur <> "" And fsco.Chemin <> "" Then
            nom = Trim(fsco.nom)
            If nom <> "" Then
                If fsco.f_attr = 1 Or fsco.f_attr = 33 Then
                    message = "Fichier en lecture seule."
                    reponse = MsgBox(message, , "Saisie du nom du fichier OpenOffice ")
                    nom = ""
        '            Text1.Text = ""
        '                Text2.Text = ""
                Else
                    Text3.Text = fsco.lecteur + fsco.Chemin + "\"
        '                Label1.Caption = Trim(fsco.nomcomplet)
        '                Label1.Caption = fsco.nomcomplet
                    Text4.Text = Left$(fsco.nom, Len(fsco.nom) - 4)
                  sav_word = False
             End If
                
            Else
        '        Text1.Text = ""
        '            Text2.Text = ""
            End If
        Else
        '       nom = ""
                Text1.Text = ""
        '            Text2.Text = ""
        End If
        End If
    Wend
Exit Sub
errHandler:
Exit Sub

End Sub

Private Sub Cmd_valid_Click()
Dim frmf As Frm_savword
Set frmf = New Frm_savword
chemin_fiche = Text3.Text + Text4.Text
sav_word = True
            If Dir(chemin_fiche) <> "" Then
                 frmf.Caption = "Sauvegarde sous OpenOffice"
                 frmf.Cmd_complet.Visible = False
                 frmf.Text2.Text = " le fichier " + Text2.Text + " existe déjà."
                 frmf.Show 1
                 sav_word = frmf.sav_w
                 mod_save = frmf.mod_sav
            Else
                mod_save = ""
                sav_word = False
            End If
    If Trim(chemin_fiche) <> "" And Not sav_word Then
        nomooo = chemin_fiche
        Cmd_ok.Enabled = True
        Call Cmd_ok_Click
    Else
        If nomooo = "" Then
        Cmd_ok.Enabled = False
        End If
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Opt_word Then
    On Error GoTo test_Error
    If Not awd Is Nothing Then
        ad.Close
        awd.Quit
       Set ad = Nothing
       Set awd = Nothing
    End If
End If
Exit Sub
test_Error:
       Set ad = Nothing
       Set awd = Nothing

End Sub

Private Sub Form_Load()
Centre Me
nomword = ""
nomooo = ""
Call Opt_imp_Click
If Not ok_wor Then
    Opt_word.Visible = False
End If
Me.Text2.Text = nomword
If Printers.count > 0 Then
Me.Label2.Caption = Printer.DeviceName
Else
Me.Label2.Caption = "Pas d'imprimante déclarée"
End If
Set owner = MDIFrm_menu.rec_owner
nom_fiche = Me.Type1 + "_" + Me.nomobjet
'If Type1 = "versant" Then
'If Opt_word Then
'    If Not owner.fbassin Is Nothing Then
'        owner.fbassin.Enabled = False
'    Else
'        owner.fobjet.Enabled = False
'    End If
'End If
End Sub
Private Sub Cmd_recfic_Click()
Dim reponse As Integer
Dim message As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim f As File
Dim d As Drive
Dim nom As String
Dim frmf As Frm_savword
Set frmf = New Frm_savword
sav_word = True
'    Set frmf = Nothing
    
    While sav_word
        cdlg1.DialogTitle = "Enregistrer sous " '"Recherche d'un fichier "
        cdlg1.FileName = ""
        cdlg1.Filter = "Fichiers WORD (*.doc)|*.doc|Tous (*.*)|*.*"
        cdlg1.InitDir = ""
        cdlg1.Flags = cdlOFNHideReadOnly
        cdlg1.ShowSave
'        Debug.Print cdlg1.Tag
 '     cdlg1.ShowOpen
        s = cdlg1.FileName
        fsco = create_fs(s)
        If fsco.dr_type = 1 Then
            message = "Fichier sur disquette;" + Chr(13) + Chr(10) + "Vérifier que la disquette n'est pas protégée en écriture."
            reponse = MsgBox(message, , "Saisie du nom du fichier WORD ")
        End If
        If fsco.dr_type = 4 Then
            message = "Fichier sur CR-ROM;" + Chr(13) + Chr(10) + "Pas d'accés en écriture."
            reponse = MsgBox(message, , "Saisie du nom du fichier WORD ")
            nom = ""
            Text1.Text = ""
        '        Text2.Text = ""
        ElseIf fsco.lecteur <> "" And fsco.Chemin <> "" Then
            nom = Trim(fsco.nom)
            If nom <> "" Then
                If fsco.f_attr = 1 Or fsco.f_attr = 33 Then
                    message = "Fichier en lecture seule."
                    reponse = MsgBox(message, , "Saisie du nom du fichier WORD ")
                    nom = ""
                    Text1.Text = ""
        '                Text2.Text = ""
                Else
                    Text1.Text = fsco.nomcomplet
        '                Label1.Caption = Trim(fsco.nomcomplet)
        '                Label1.Caption = fsco.nomcomplet
                    Text2.Text = fsco.nomcomplet
               End If
                
            Else
                Text1.Text = ""
        '            Text2.Text = ""
            End If
        Else
                nom = ""
                Text1.Text = ""
        '            Text2.Text = ""
        End If
        If nom <> "" Then
            If Dir(nom) <> "" Then
        '            me.Enabled=false
        '             frmf.Label1.Caption = " le fichier " + Label1.Caption + " existe déjà."
                 frmf.Caption = "Sauvegarde sous WORD"
                 frmf.Text2.Text = " le fichier " + Text2.Text + " existe déjà."
                 frmf.Show 1
                 sav_word = frmf.sav_w
                 mod_save = frmf.mod_sav
            Else
                mod_save = ""
                sav_word = False
            End If
        Else
            sav_word = False
        End If
    Wend
    If Trim(nom) <> "" And Not sav_word Then
        nomword = Text2.Text
        Cmd_ok.Enabled = True
    Else
        If nomword = "" Then
        Cmd_ok.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If Opt_word Then
'    If Type1 = "versant" Then
'    If Not owner.fbassin Is Nothing Then
'        owner.fbassin.Enabled = True
'    Else
'        owner.fobjet.Enabled = True
'    End If
'End If
End Sub

Private Sub Opt_imp_Click()
    Frame3.Visible = True  'imprimante
    Label2.Visible = True
    Cmd_config.Visible = True
    
    Frame2.Visible = False  'Word
    Text1.Visible = False
    Text2.Visible = False
    Cmd_recfic.Visible = False
    
    Frame4.Visible = False  'OOO
    Text3.Visible = False
    Text4.Visible = False
    Cmd_repos.Visible = False
    Cmd_valid.Visible = False
    
    Cmd_ok.Caption = "Aperçu"   'commun
    Cmd_ok.Enabled = True
    Label3.Caption = ""
End Sub

Private Sub Opt_OOO_Click()
    Frame3.Visible = False  'imprimante
    Label2.Visible = False
    Cmd_config.Visible = False
    
    Frame2.Visible = False  'Word
    Text1.Visible = False
    Text2.Visible = False
    Cmd_recfic.Visible = False
    
    Frame4.Visible = True  'OOO
    Text3.Visible = True
    Text3.Text = chemin_etude
    Text4.Visible = True
    Text4.Text = nom_fiche + ".odt"
    Cmd_repos.Visible = True
    Cmd_valid.Visible = True
    Cmd_valid.SetFocus
    
    Cmd_ok.Caption = "OK" 'commun
    If nomooo = "" Then
        Cmd_ok.Enabled = False
    Else
        Cmd_ok.Enabled = True
    End If
    Label3.Caption = ""
'    owner.Enabled = False

End Sub

Private Sub Opt_word_Click()
    Frame3.Visible = False  'imprimante
    Label2.Visible = False
    Cmd_config.Visible = False
    
    Frame4.Visible = False  'OOO
    Text3.Visible = False
    Text4.Visible = False
    Cmd_repos.Visible = False
    Cmd_valid.Visible = False

    Frame2.Visible = True   'Word
    Text1.Visible = False
    Text2.Visible = True
    Cmd_recfic.Visible = True
    Cmd_recfic.SetFocus
    
    Cmd_ok.Caption = "OK"  'commun
    If nomword = "" Then
        Cmd_ok.Enabled = False
    Else
        Cmd_ok.Enabled = True
    End If
    Label3.Caption = ""
'    owner.Enabled = False
End Sub
