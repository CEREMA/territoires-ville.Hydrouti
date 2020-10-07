Attribute VB_Name = "Fonctions_OOO"
Function AddODT(ByVal nom As String, ByRef aNomfich() As String) As Boolean
Dim cZip As ccZip
Dim i As Integer, lRet As Integer
 AddODT = False
 If UBound(aNomfich) < 1 Then Exit Function
  
 ' If sPath = "" Then Exit Sub
  
  If cZip Is Nothing And nom <> "" Then
    Set cZip = New ccZip
    cZip.Init nom
  ElseIf cZip Is Nothing Then
    Exit Function
  End If
  cZip.Comm = "" ' txtComm
  cZip.Level = CLng(txtLevel)
'  lRet = cZip.AddFile(sPath, , False)
For i = 1 To UBound(aNomfich)
 lRet = cZip.AddFile(aNomfich(i), , False)
  If lRet = 0 Then AddODT = True
Next
' lRet = cZip.AddFile("C:\Pictures", True, False)
'  If lRet = 0 Then AddODT = True

End Function
Public Sub del_fichiers(ByRef aNomfich() As String)
Dim i As Integer
For i = 1 To UBound(aNomfich)
 Kill aNomfich(i)
Next

End Sub


Public Sub del_fich_rep(ByVal nom As String)
Dim myfile As String
ChDir (nom)
myfile = Dir("*.*")
While myfile <> ""
    Kill myfile
    myfile = Dir("*.*")
Wend
ChDir (chemin_app)
End Sub

Public Sub Ecr_debut()
Dim nom As String, chaine As String
Dim ok As Boolean, ok1 As Boolean
    lhFicooo = FreeFile
ok = False
ok1 = False
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_debut_content" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_debut_content" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhFicooo1, chaine
        End If

    Loop
    Close #lhFicooo
End Sub
Public Sub Ecr_styles(ByVal nomfich As String, ByVal xdate As String)
Dim nom As String, chaine As String
Dim lhfic As Integer
Dim ok As Boolean, ok1 As Boolean
Dim sep As String
sep = Chr(34)
ok = False
ok1 = False
   lhfic = FreeFile
    Open nomfich For Output As #lhfic
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_debut_styles" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_debut_styles" Then
            ok1 = True
        End If
        If ok1 And ok Then
            If Trim(chaine) = "service1" Then
'chaine = "<table:table-cell table:style-name=" & sep & c_style_cel & sep & " office:value-type=" & sep & "string" & sep & ">"
'chaine = "<text:span text:style-name=" & sep &"T2"& sep & ">"Centre d'études Techniques de l'Equipement de l'Est& "</text:span>"
chaine = "<text:span text:style-name=" & sep & "T2" & sep & ">" & text_serv1 & "</text:span>"
            End If
            If Trim(chaine) = "service2" Then
chaine = "<text:span text:style-name=" & sep & "T2" & sep & ">" & text_serv2 & "</text:span>"
            End If
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
    
    Print #lhfic, "dossier : " + nom_fich_edit
    
    chaine = "</text:p>"
    Print #lhfic, chaine
    chaine = "<text:p text:style-name=" & sep & "P7" & sep & ">"
    Print #lhfic, chaine
    chaine = "<text:tab />"
    Print #lhfic, chaine

    Print #lhfic, xdate
ok = False
ok1 = False
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
         If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_fin_styles" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_fin_styles" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
    Close #lhfic

End Sub
Public Sub Ecr_meta(ByVal nomfich As String, ByVal xdate As String)
Dim nom As String, chaine As String
Dim lhfic As Integer
Dim ok As Boolean, ok1 As Boolean
ok = False
ok1 = False
    lhfic = FreeFile
    Open nomfich For Output As #lhfic
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_debut_meta" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_debut_meta" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
'    chaine = "<meta:initial-creator>xxx</meta:initial-creator>"
    chaine = "<meta:initial-creator></meta:initial-creator>"
    Print #lhfic, chaine
    chaine = "<meta:creation-date>" & xdate & "</meta:creation-date>"
    Print #lhfic, chaine
'    chaine = "<dc:creator>xxx</dc:creator>"
    chaine = "<dc:creator></dc:creator>"
    Print #lhfic, chaine
    chaine = "<dc:date>" & xdate & "</dc:date>"
    Print #lhfic, chaine
'    chaine = "<meta:printed-by>xxx</meta:printed-by>"
    chaine = "<meta:printed-by></meta:printed-by>"
    Print #lhfic, chaine
    chaine = "<meta:print-date>" & xdate & "</meta:print-date>"
    Print #lhfic, chaine

ok = False
ok1 = False
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_fin_meta" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_fin_meta" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
    Close #lhfic

End Sub

Public Sub Ecr_manifest(ByVal nomfich As String, ByVal Image1 As String, ByVal Image2 As String)
Dim nom As String, chaine As String, sep As String
Dim lhfic As Integer
Dim ok As Boolean, ok1 As Boolean
ok = False
ok1 = False
sep = Chr(34)
    lhfic = FreeFile
    Open nomfich For Output As #lhfic
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_debut_manifest" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_debut_manifest" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
    chaine = "<manifest:file-entry manifest:media-type=" & sep & "image/bmp" & sep & " manifest:full-path=" & sep & Image1 & sep & "/>"
    Print #lhfic, chaine
    If Image2 <> "" Then
        chaine = "<manifest:file-entry manifest:media-type=" & sep & "image/bmp" & sep & " manifest:full-path=" & sep & Image2 & sep & "/>"
        Print #lhfic, chaine
    End If
ok = False
ok1 = False
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_fin_manifest" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_fin_manifest" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhfic, chaine
        End If
    Loop
    Close #lhFicooo
    Close #lhfic

End Sub

Public Sub recopy_debut(ByVal nom As String)
Dim chaine As String
Dim ok As Boolean
Dim sep As String
sep = Chr(34)
   lhFicooo = FreeFile
    ok = True
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If Trim(chaine) = "</office:text>" Then
            ok = False
        End If
        If ok Then
            Print #lhFicooo1, chaine
        End If
    Loop
    Close #lhFicooo
    chaine = "<text:p text:style-name=" & sep & "P9" & sep & " />"
    Print #lhFicooo1, chaine

End Sub

Public Sub Ecr_fin()
Dim nom As String, chaine As String
Dim ok As Boolean, ok1 As Boolean
ok = False
ok1 = False
    lhFicooo = FreeFile
    nom = chemin_app + "ini_xml.txt"
    Open nom For Input As #lhFicooo
    Do While Not EOF(lhFicooo)
        Input #lhFicooo, chaine
        If ok1 And Not ok Then
            ok = True
        End If
        If Trim(chaine) = "fin_fin_content" Then
            ok1 = False
        End If
        If Trim(chaine) = "deb_fin_content" Then
            ok1 = True
        End If
        If ok1 And ok Then
            Print #lhFicooo1, chaine
        End If
    Loop
    Close #lhFicooo
End Sub
Public Sub Ecr_titre(ByVal nom1 As String, ByVal nom2 As String, ByVal noPage As Integer)
Dim chaine As String, nometude As String
Dim sep As String
Dim i As Integer
Dim xdate As String
xdate = Date
sep = Chr(34)
''*****ecrit date
'    index = index + 1
'    chaine = "<draw:frame text:anchor-type=" & sep & "page" & sep & " text:anchor-page-number=" & sep
'    chaine = chaine & LTrim$(Str$(noPage)) & sep & "  draw:z-index=" & sep & LTrim$(Str$(index)) & sep & " draw:style-name="
' '   chaine = chaine & sep & "gr2" & sep & " draw:text-style-name=" & sep & "P1" & sep & "  svg:width=" & sep & "3.387cm" & sep & "  svg:height=" & sep & "0.468cm" & sep & "  svg:x=" & sep & "2.385cm" & sep & "  svg:y=" & sep & "26.903cm" & sep & " >"
'    chaine = chaine & sep & "gr2" & sep & " draw:text-style-name=" & sep & "P1" & sep & "  svg:width=" & sep & "3.387cm" & sep & "  svg:height=" & sep & "0.468cm" & sep & "  svg:x=" & sep & "3.39cm" & sep & "  svg:y=" & sep & "27.058cm" & sep & " >"
'    Print #lhFicooo1, chaine
'    chaine = "<draw:text-box>"
'    Print #lhFicooo1, chaine
'    chaine = "<text:p text:style-name=" & sep & "P11" & sep & ">"
'    Print #lhFicooo1, chaine
'    chaine = "<text:span text:style-name=" & sep & "T11" & sep & ">" & xdate & "</text:span>"
'     Print #lhFicooo1, chaine
'    chaine = "</text:p>"
'    Print #lhFicooo1, chaine
'    chaine = "</draw:text-box>"
'    Print #lhFicooo1, chaine
'    chaine = "</draw:frame>"
'    Print #lhFicooo1, chaine
'    index = index + 1
'    chaine = "<draw:frame text:anchor-type=" & sep & "page" & sep & " text:anchor-page-number=" & sep
'    chaine = chaine & LTrim$(Str$(noPage)) & sep & "  draw:z-index=" & sep & LTrim$(Str$(index)) & sep & " draw:style-name="
' '   chaine = chaine & sep & "gr2" & sep & " draw:text-style-name=" & sep & "P1" & sep & "  svg:width=" & sep & "3.387cm" & sep & "  svg:height=" & sep & "0.468cm" & sep & "  svg:x=" & sep & "2.385cm" & sep & "  svg:y=" & sep & "26.903cm" & sep & " >"
'    chaine = chaine & sep & "gr2" & sep & " draw:text-style-name=" & sep & "P1" & sep & "  svg:width=" & sep & "0.953cm" & sep & "  svg:height=" & sep & "0.532cm" & sep & "  svg:x=" & sep & "16.646cm" & sep & "  svg:y=" & sep & "27.058cm" & sep & " >"
'    Print #lhFicooo1, chaine
'    chaine = "<draw:text-box>"
'    Print #lhFicooo1, chaine
'    chaine = "<text:p text:style-name=" & sep & "P12" & sep & ">"
'    Print #lhFicooo1, chaine
'    chaine = "<text:span text:style-name=" & sep & "T11" & sep & ">" & LTrim$(Str$(noPage)) & "</text:span>"
'     Print #lhFicooo1, chaine
'    chaine = "</text:p>"
'    Print #lhFicooo1, chaine
'    chaine = "</draw:text-box>"
'    Print #lhFicooo1, chaine
'    chaine = "</draw:frame>"
'    Print #lhFicooo1, chaine
'
'
'
'
'******ecrit titres
    Index = Index + 1
    nometude = "Etude : " + nom_etude
chaine = "<text:p>"
    Print #lhFicooo1, chaine
chaine = "<text:s text:c=" & sep & "20" & sep & " />"
    Print #lhFicooo1, chaine

chaine = "<text:span text:style-name=" & sep & "T2" & sep & ">" & nometude & "</text:span>"
    Print #lhFicooo1, chaine
chaine = "</text:p>"
    Print #lhFicooo1, chaine
    chaine = "<text:p text:style-name=" & sep & "Standard" & sep & ">"
    Print #lhFicooo1, chaine
    chaine = "<draw:frame text:anchor-type=" & sep & "paragraph" & sep & " draw:z-index="
    chaine = chaine & sep & LTrim$(str$(Index)) & sep & " draw:style-name=" & sep & "gr1"
    chaine = chaine & sep & " draw:text-style-name=" & sep & "P2" & sep & " svg:width="
    chaine = chaine & sep & "14.368cm" & sep & " svg:height=" & sep & "1.086cm"
    chaine = chaine & sep & " svg:x=" & sep & "1.36cm" & sep & " svg:y=" & sep & "0.325cm" & sep & ">"
    Print #lhFicooo1, chaine
    chaine = "<draw:text-box>"
    Print #lhFicooo1, chaine
    chaine = "<text:p text:style-name=" & sep & "P1" & sep & ">"
    Print #lhFicooo1, chaine
    chaine = "<text:span text:style-name=" & sep & "T1" & sep & ">" & nom1 & "</text:span>"
    Print #lhFicooo1, chaine
    chaine = "</text:p>"
    Print #lhFicooo1, chaine
    chaine = "<text:p text:style-name=" & sep & "P1" & sep & ">"
    Print #lhFicooo1, chaine
    chaine = "<text:span text:style-name=" & sep & "T2" & sep & ">" & nom2 & "</text:span>"
    Print #lhFicooo1, chaine
    chaine = "</text:p>"
    Print #lhFicooo1, chaine
    chaine = "</draw:text-box>"
    Print #lhFicooo1, chaine
    chaine = "</draw:frame>"
    Print #lhFicooo1, chaine
    chaine = "</text:p>"
    Print #lhFicooo1, chaine
    For i = 1 To 3
        chaine = "<text:p text:style-name=" & sep & "Standard" & sep & " />"
        Print #lhFicooo1, chaine
    Next

End Sub

Public Sub Ecr_titre_tableau(ByVal nb_av As Integer, ByVal nom1 As String)
Dim chaine As String, sep As String
Dim i As Integer
sep = Chr(34)
    For i = 1 To nb_av
        chaine = "<text:p text:style-name=" & sep & "Standard" & sep & " />"
        Print #lhFicooo1, chaine
    Next
    chaine = "<text:p text:style-name=" & sep & "P3" & sep & ">" & nom1 & "</text:p>"
    Print #lhFicooo1, chaine
End Sub
    
Public Sub Ecr_tableau(ByVal stitre As String, ByRef liste1() As Variant, ByRef liste2() As String)
Dim chaine As String
Dim sep As String, deb_style_cel As String, style_cel As String
Dim i As Integer, j As Integer
sep = Chr(34)
    chaine = "<table:table table:name=" & sep & "Tableau" & stitre & sep & " table:style-name=" & sep & "Tableau1" & sep & ">"
    Print #lhFicooo1, chaine
    For j = 1 To UBound(liste2)
    chaine = "<table:table-column table:style-name=" & sep & "Tableau1." & liste2(j, 1) & sep & " />"
    Print #lhFicooo1, chaine
    Next
'    chaine = "<table:table-header-rows>"
'    Print #lhFicooo1, chaine
'    chaine = "<table:table-row>"
'    Print #lhFicooo1, chaine
    For i = 0 To UBound(liste1)
        chaine = "<table:table-row>"
        Print #lhFicooo1, chaine
        If i = 0 Then
            deb_style_cel = "H"
        ElseIf i = UBound(liste1) Then
            deb_style_cel = "B"
        Else
             deb_style_cel = "M"
        End If
        For j = 1 To UBound(liste2)
            If j = 1 Then
                style_cel = deb_style_cel + "G"
                If UBound(liste1) = 0 Then
                    style_cel = "TG"
                End If
            ElseIf j = UBound(liste2) Then
                style_cel = deb_style_cel + "D"
                If UBound(liste1) = 0 Then
                    style_cel = "TD"
                End If
            Else
                 style_cel = deb_style_cel + "M"
                 If UBound(liste1) = 0 Then
                    style_cel = "TM"
                End If
           End If
            Call Ecr_lign(liste1(i, j), style_cel, liste2(j, 2))
        Next
        chaine = "</table:table-row>"
        Print #lhFicooo1, chaine
    Next
        chaine = "</table:table>"
        Print #lhFicooo1, chaine
 
End Sub
Public Sub Ecr_lign(ByVal c_texte As String, ByVal c_style_cel As String, ByVal c_style_texte As String)
Dim chaine As String, sep As String
sep = Chr(34)
    chaine = "<table:table-cell table:style-name=" & sep & c_style_cel & sep & " office:value-type=" & sep & "string" & sep & ">"
    Print #lhFicooo1, chaine
    chaine = "<text:p text:style-name=" & sep & c_style_texte & sep & " >" & c_texte & "</text:p>"
    Print #lhFicooo1, chaine
    chaine = "</table:table-cell>"
    Print #lhFicooo1, chaine

End Sub
Public Sub Ecr_dess(ByVal nb_av As Integer, ByVal haut As String, ByVal nom As String)
Dim chaine As String, sep As String
sep = Chr(34)
Index = Index + 1
    For i = 1 To nb_av
        chaine = "<text:p text:style-name=" & sep & "Standard" & sep & " />"
        Print #lhFicooo1, chaine
    Next
    chaine = "<text:p text:style-name=" & sep & "P8" & sep & " >"
    Print #lhFicooo1, chaine

    chaine = "<draw:frame draw:style-name=" & sep & "fr1" & sep & " draw:name=" & sep & "Image2" & sep
    chaine = chaine & " text:anchor-type=" & sep & "paragraph" & sep & " svg:x=" & sep & "1.281cm" & sep
    chaine = chaine & " svg:y=" & sep & "0.499cm" & sep & " svg:width=" & sep & "14.42cm" & sep
    chaine = chaine & " svg:height=" & sep & haut & "cm" & sep & " draw:z-index=" & sep & LTrim$(str$(Index)) & sep & " >"
    Print #lhFicooo1, chaine
    chaine = "<draw:image xlink:href=" & sep & nom & sep & " xlink:type=" & sep
    chaine = chaine & "simple" & sep & " xlink:show=" & sep & "embed" & sep & " xlink:actuate=" & sep & "onLoad" & sep & " />"
    Print #lhFicooo1, chaine
    chaine = "</draw:frame>"
    Print #lhFicooo1, chaine
    chaine = "</text:p>"
    Print #lhFicooo1, chaine
 End Sub

