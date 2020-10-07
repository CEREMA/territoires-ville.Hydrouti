Attribute VB_Name = "Fonctions"
Function calc_date(ByVal nom As String) As String
Dim nomb1 As String
Dim nannee As String, nmois As String, njour As String
Dim nheure As String, nminute As String, nseconde As String
nomb1 = ""
'Debug.Print nom
nannee = Mid(nom, 7, 4)
nmois = Mid(nom, 4, 2)
njour = Mid(nom, 1, 2)
nheure = Mid(nom, 12, 2)
nminute = Mid(nom, 15, 2)
nseconde = Mid(nom, 18, 2)
nomb1 = Trim(nannee) + Trim(nmois) + Trim(njour) + Trim(nheure) + Trim(nminute) + Trim(nseconde)
'Debug.Print nomb1
'Debug.Print Val(nomb1)
calc_date = nomb1
End Function
Public Function recup_defchamp(ByVal nom_frm As String, ByVal nom_chp As String, ByVal num_ind As Integer, _
    ByVal nvalue As Double) As Boolean
Dim wk As Workspace
Dim rcs1 As Recordset
Dim dbs1 As Database
Dim nom As String, sreq As String, mes1 As String, munit As String
Dim defc As defchamp
Dim reponse As Integer
Dim ok As Boolean
ok = True
' chgt defchamp en dbf
'Set wk = DBEngine.Workspaces(0)
'Set dbs1 = wk.OpenDatabase("C:\Hydraulique\bo_v800_600", False, True, "dBASE III;")
'    sreq = "select *    from defchamps    where Form = '" & nom_frm & "' and Nomchp = '" & nom_chp
'''defchamps.hyo est une base ACCESS
    nom = chemin_app + "defchamps.hyo"
     Set dbs1 = OpenDatabase(nom)
    sreq = "select *    from defchamps    where Form = '" & nom_frm & "' and Nomchp = '" & nom_chp
    If num_ind >= 0 Then
        sreq = sreq & "' and Indexc = " & num_ind & ";"
    Else
        sreq = sreq & "'  ;"
    End If
     Set rcs1 = dbs1.OpenRecordset(sreq)
    If rcs1.RecordCount > 0 Then
        With rcs1
            mes1 = ""
            If Trim(rcs1.Fields("message")) <> "" Then
                mes1 = rcs1.Fields("message")
            End If
            munit = ""
            If Trim(rcs1.Fields("unite")) <> "" Then
                munit = rcs1.Fields("unite")
            End If
            With defc
                defc.Form = rcs1.Fields("form")
                defc.Intitule = rcs1.Fields("intitule")
                defc.Ancnom = rcs1.Fields("ancnom")
                defc.Nomchp = rcs1.Fields("nomchp")
                defc.Indexc = rcs1.Fields("indexc")
                defc.taille = rcs1.Fields("taille")
                defc.Decimal = rcs1.Fields("decimal")
                defc.OKmini = rcs1.Fields("okmini")
                defc.Mini = rcs1.Fields("mini")
                defc.OKmaxi = rcs1.Fields("okmaxi")
                defc.Maxi = rcs1.Fields("maxi")
                defc.message = mes1
                defc.Chplabel = rcs1.Fields("chplabel")
                defc.Label = rcs1.Fields("label")
                defc.Chpunite = rcs1.Fields("chpunite")
                defc.Unite = munit
            End With
        End With
    End If
rcs1.Close
dbs1.Close
If Trim(defc.message) <> "" Then
    If Trim(defc.Form) = "Frm_do" Then
        defc.Label = "Longueur de la canalisation"
        If Trim(defc.Nomchp) = "Tb_amo" And defc.Indexc = 3 Then
            defc.Maxi = edessdo.lgdisp
        End If
         If Trim(defc.Nomchp) = "Tb_ava" And defc.Indexc = 3 Then
            defc.Maxi = edessdo.lgdisp - edessdo.Lam
        End If
    End If
    ok = verif_champ(defc, nvalue)
End If
recup_defchamp = ok
End Function
Public Function verif_champ(ByRef defc As defchamp, nvalue As Double) As Boolean
Dim ok As Boolean
Dim mes_verif As String
ok = True
    If defc.OKmini And defc.OKmaxi Then
        If nvalue < defc.Mini Or nvalue > defc.Maxi Then
'            reponse = MsgBox(defc.Label + " : " + defc.message, , "Validation de la valeur")
                mes_verif = defc.message + " :" + str$(defc.Mini) + " |" + str$(defc.Maxi)
'            reponse = MsgBox(defc.message, , defc.Label)
            reponse = MsgBox(mes_verif, , defc.Label)
            ok = False
        End If
    End If
    If Not defc.OKmini And defc.OKmaxi Then
        If nvalue > defc.Maxi Then
'            reponse = MsgBox(defc.Label + " : " + defc.message, , "Validation de la valeur")
            mes_verif = defc.message + " :" + str$(defc.Maxi)
'            reponse = MsgBox(defc.message, , defc.Label)
            reponse = MsgBox(mes_verif, , defc.Label)
            ok = False
        End If
    End If
    If defc.OKmini And Not defc.OKmaxi Then
        If nvalue < defc.Mini Then
'            reponse = MsgBox(defc.Label + " : " + defc.message, , "Validation de la valeur")
            mes_verif = defc.message + ":" + str$(defc.Mini)
            reponse = MsgBox(defc.message, , defc.Label)
            ok = False
        End If
    End If
    verif_champ = ok
End Function
Public Function Virgule(ByVal ValNum As Double, Precision As Integer) As String
'remplacement de la virgule par un point et formatage des valeurs entrées à la précision voulue
If Precision = 1 Then
    Virgule = Format(ValNum, "###0.0")
ElseIf Precision = 2 Then
    Virgule = Format(ValNum, "###0.00")
ElseIf Precision = 3 Then
    Virgule = Format(ValNum, "###0.000")
ElseIf Precision = 4 Then
    Virgule = Format(ValNum, "###0.0000")
Else
    Virgule = Format(ValNum, "###0")
End If
End Function
Public Function rempl_virgule(ByVal chaine As String) As String
'remplacement de la virgule par un point dans une chaine
    Dim j As Integer
    j = InStr(1, chaine, ",")
    If j <> 0 Then    'séparateur virgule
        Mid(chaine, j, 1) = "."
    End If
    If j = 1 Then
        chaine = "0" + chaine
    End If
    rempl_virgule = chaine
End Function
Public Function ajout_zero(ByVal chaine As String) As String
    Dim j As Integer
    j = InStr(1, chaine, ".")
    If j = 1 Then
        chaine = "0" + chaine
    End If
    ajout_zero = chaine
End Function
Public Function txtVersNum(txt As String) As Double
Dim PosPoint As Long
'remplacement de la virgule par un point pour recup de la valeur
PosPoint = InStr(1, txt, ",")
If PosPoint <> 0 Then    'séparateur virgule
    Mid(txt, PosPoint, 1) = "."
End If
txtVersNum = val(txt)
End Function

Public Function verif_car(ByVal nom As String, ByVal KeyAscii, ByVal mes As String, ByVal ctyp As String) As Integer
Dim reponse As Integer
Dim ok As Boolean
verif_car = KeyAscii
 Select Case ctyp
    Case Is = "R"
        ok = rech_point(nom)
        If (KeyAscii = 46 And ok) Or (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13 Then
            reponse = MsgBox("Caractére non valide", , mes)
            verif_car = 0
        End If
    Case Is = "I"
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
            reponse = MsgBox("Caractére non valide", , mes)
            verif_car = 0
        End If
End Select
End Function
Public Function verif_cart0(ByVal nom As String, ByVal mes As String, ByVal ctyp As String) As String
Dim reponse As Integer
Dim ok As Boolean
Dim nom0 As String
Dim scar As String
Dim i As Integer
reponse = 1
verif_cart0 = ""
nom0 = ""
i = 1
While i <= Len(nom) And reponse > 0
    scar = Mid$(nom, i, 1)
    reponse = verif_cart(nom0, Asc(scar), mes, ctyp)
    If reponse > 0 Then
        nom0 = nom0 + scar
    End If
    i = i + 1
Wend
If reponse > 0 Then
    verif_cart0 = "ok"
End If
End Function
Public Function verif_cart(ByVal nom As String, ByVal KeyAscii As Integer, ByVal mes As String, ByVal ctyp As String) As Integer
Dim reponse As Integer
Dim ok As Boolean
verif_cart = KeyAscii
 Select Case ctyp
    Case Is = "R"
        ok = rech_point(nom)
        If (KeyAscii = 46 And ok) Or (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 13 Then
            reponse = MsgBox("Caractére non valide", , mes)
            verif_cart = 0
        End If
    Case Is = "I"
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
            reponse = MsgBox("Caractére non valide", , mes)
            verif_cart = 0
        End If
End Select
End Function
Public Sub key13b(ByRef Fm As Form)
Dim itab As Integer, icotab As Integer
itab = Fm.ActiveControl.TabIndex
icotab = rech_tab_suib(Fm, itab)
Fm.Controls(icotab).SetFocus
End Sub
Public Function rech_list(ByRef l_tb() As Variant, ByVal stab As String) As Variant
Dim i1 As Integer, j1 As Integer, ok As Boolean
Dim l() As Variant
If UBound(l_tb) > 0 Then
ok = False
For i1 = 1 To UBound(l_tb)
    For j1 = 1 To UBound(l_tb(i1))
'        Debug.Print i1, j1, l_tb(i1)(j1)
        If UCase$(stab) = UCase$(l_tb(i1)(j1)) Then
            ok = True
'            Debug.Print i1, j1
            Exit For
        End If
    Next
    If ok Then
        Exit For
    End If
Next
If i1 > UBound(l_tb) Then
    l = l_tb(0)
    Else
    l = l_tb(i1)
End If
Else
l = l_tb(0)
End If
rech_list = l


End Function
Public Sub donne_focus(ByRef Fm As Form)
 Dim itab As Integer
' For i = 1 To Fm.Controls.Count
'    If Fm.Controls(i - 1).TabIndex = Fm.Tb_amo(0).TabIndex Then
'    itab = i - 1
'    i = Fm.Controls.Count
'    End If
'Next
' Fm.Controls(itab).SetFocus
'
End Sub

Public Sub key13(ByRef Fm As Object)
Dim itab As Integer, icotab As Integer
Dim l() As Variant, i1 As Integer, stab As String
l = Fm.get_l_tb()
itab = Fm.ActiveControl.TabIndex
stab = Fm.ActiveControl.Name
l = rech_list(l, stab)
'For i1 = 0 To UBound(l)
' Debug.Print i1, l(i1)
'Next


icotab = rech_tab_sui(Fm, itab, l)
Fm.Controls(icotab).SetFocus
End Sub
Function rech_tab_suib(ByRef Fm As Form, ByVal idx As Integer) As Integer
Dim ok As Boolean
Dim idc As Integer, idcmin As Integer
Dim i As Integer
ok = False
While Not ok
idcmin = 1000
ok = False
 For i = 1 To Fm.Controls.count
    If verif_tab(Fm.Controls(i - 1)) Then
    If Fm.Controls(i - 1).TabStop Then
'     Debug.Print i; Me.Controls(i - 1).Name; Me.Controls(i - 1).TabIndex
'     Debug.Print i, Me.Controls(i - 1).TabIndex
        If Fm.Controls(i - 1).TabIndex > idx Then
            If Fm.Controls(i - 1).Enabled And Fm.Controls(i - 1).Visible Then
            If Fm.Controls(i - 1).TabIndex < idcmin Then
                ok = True
                idcmin = Fm.Controls(i - 1).TabIndex
                idc = i - 1
            End If
            End If
        End If
    End If
    End If
  Next
If Not ok Then
    idx = -1
End If
Wend
rech_tab_suib = idc
End Function
Function rech_tab_sui(ByRef Fm As Form, ByVal idx As Integer, _
    ByRef l() As Variant) As Integer
Dim ok As Boolean
Dim idc As Integer, idcmin As Integer
Dim i As Integer, i1 As Integer
ok = False
While Not ok
idcmin = 1000
ok = False
 For i = 1 To Fm.Controls.count
     If verif_tab(Fm.Controls(i - 1)) Then
     If Fm.Controls(i - 1).TabStop Then
'      Debug.Print i; Fm.Controls(i - 1).Name; Fm.Controls(i - 1).TabStop; Fm.Controls(i - 1).TabIndex
'      Debug.Print i, Fm.Controls(i - 1).TabIndex
       If verif_l_tab(l, Fm.Controls(i - 1).Name) Then
        If Fm.Controls(i - 1).TabIndex > idx Then
            If Fm.Controls(i - 1).Enabled And Fm.Controls(i - 1).Visible Then
                If Fm.Controls(i - 1).TabIndex < idcmin Then
                    ok = True
                    idcmin = Fm.Controls(i - 1).TabIndex
                    idc = i - 1
                End If
            End If
        End If
        End If
    End If
    End If
  Next
If Not ok Then
    idx = -1
End If
Wend
rech_tab_sui = idc
End Function
Public Function verif_l_tab(ByRef l() As Variant, ByVal stab As String) As Boolean
Dim ok As Boolean, i As Integer
ok = False
If UBound(l) = 0 Then
    ok = True
Else
For i = 1 To UBound(l)
    If UCase$(stab) = UCase$(l(i)) Then
        ok = True
        i = UBound(l)
    End If
Next
End If
 verif_l_tab = ok
End Function

Function verif_tab(ByRef co As Control) As Boolean
Dim ok As Boolean
On Error GoTo erreur
ok = co.TabStop
verif_tab = True
Exit Function
erreur:
verif_tab = False
End Function
Public Function rech_point(ByVal nom As String) As Boolean
Dim i As Integer
rech_point = False
For i = 1 To Len(nom)
    If Mid(nom, i, 1) = "." Or Mid(nom, i, 1) = "," Then
        rech_point = True
    End If
Next
End Function
Public Function arccosinus(ByVal X As Double)
If X = -1 Then
    arccosinus = 4 * Atn(1)
ElseIf X = 1 Then
    arccosinus = 0
Else
    arccosinus = Atn(-X / Sqr((-X * X + 1))) + 2 * Atn(1)
End If
End Function
Public Function minimum(ByVal X As Variant, ByVal Y As Variant) As Variant
If X > Y Then
minimum = Y
Else
minimum = X
End If
End Function
Public Function maximum(ByVal X As Variant, ByVal Y As Variant) As Variant
If X > Y Then
maximum = X
Else
maximum = Y
End If
End Function

