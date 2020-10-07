Attribute VB_Name = "mod_pompe"
Function long_pompe(ByRef ebc As st_Pompe, ByRef res_am As debit_conduit, ByRef res_av As debit_conduit, ByRef hee As Double, ByRef lee As Double) As Double
'Dim dh As Double, h0 As Double, h1 As Double, h2 As Double, D0 As Double
'
'D0 = ebc.tron_amo.conduit.Diametre
'h0 = res_am.hautava
'h2 = res_av.hautamo
'dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
'long_pompe = dh
'Dim F0 As Double, he As Double
'F0 = res_am.debit / ((h0 ^ 2) * ((9.81 * D0) ^ 0.5))
'he = h0 * ((2 * F0 ^ 2) / (1 + 2 * F0 ^ 2)) ^ (2# / 3#)
'Dim le As Double
'le = h0 * (5 + 0.9 * F0)
'lee = le
'hee = he
'Dim Zsup1 As Double, Zinf As Double
'dh = dh - h2
'Zinf = -dh / h0
'dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
'dh = dh + he - h2
'ebc.h0 = dh
'Zsup1 = -dh / h0
'Zsup1 = Zinf
'' Zsup1=-1/3 X -1/4 X^2
''1/4 X^2 + 1/3 X +Zinf =0
'Dim a As Double, b As Double, c As Double, delta As Double
'a = 0.25
'b = (1 / 3)
'c = Zinf
'delta = b * b - 4 * a * c
'x1 = (-b + Sqr(delta)) / (2 * a)
'x2 = (-b - Sqr(delta)) / (2 * a)
'Dim X As Double
'X = x1 * (F0 ^ 0.8 * h0)
'' Zsup1=1 -(1/3-0.06)X -1/4 X^2
''1/4 X^2 +(1/3-0.06) X +Zsup -1 =0
''Dim a As Double, b As Double, c As Double, delta As Double
'a = 0.25
'b = (1 / 3 - 0.06 * he / h0)
'c = Zsup1 - he / h0
'delta = b * b - 4 * a * c
'x1 = (-b + Sqr(delta)) / (2 * a)
'x2 = (-b - Sqr(delta)) / (2 * a)
''Dim X As Double
'X = x1 * (F0 ^ 0.8 * h0)
'long_pompe = X
End Function
Function dess_pompe1(ByRef uc_g As UC_graphique, ByRef ebc As st_Pompe, ByRef res_am As debit_conduit, ByRef res_av As debit_conduit) As Boolean
'Dim dh As Double, h0 As Double, h1 As Double, h2 As Double, D0 As Double
'
'Dim xy(11, 2) As Double, X As Double, dx As Double, Y As Double, x0 As Double, y0 As Double
'uc_g.redef_drwidth 2
'
'D0 = ebc.tron_amo.conduit.Diametre
'h0 = res_am.hautava
'h2 = res_av.hautamo
'h1 = res_am.hautava
'
'dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
'Dim F0 As Double, he As Double
'F0 = res_am.debit / ((h0 ^ 2) * ((9.81 * D0) ^ 0.5))
'he = h0 * ((2 * F0 ^ 2) / (1 + 2 * F0 ^ 2)) ^ (2# / 3#)
'Dim le As Double
'le = h0 * (5 + 0.9 * F0)
'Dim Zsup1 As Double, Zinf As Double
'
'x0 = ebc.Long
'
'dx = x0 / 10
'For i = 1 To 11
'x1 = dx * (i - 1)
'xy(i, 1) = x1
'X = x1 / (F0 ^ 0.8 * h0)
'Y = he - (1 / 3 * h0 - 0.06 * he) * X - 0.25 * h0 * X ^ 2
'xy(i, 2) = Y
'Next
'xam = ebc.tron_amo.Absava
'yam = ebc.tron_amo.radava
'x0 = xam
'y0 = yam + he
'For i = 1 To 11
'    X = xam + xy(i, 1)
'    Y = yam + xy(i, 2)
'    uc_g.dess_lign x0, y0, X, Y, couleur.bleu, 2
'    x0 = X
'    y0 = Y
'Next
''courbe inferieure
'dh = dh - h2
'Zinf = -dh / h0
'' Zsup1=-1/3 X -1/4 X^2
''1/4 X^2 + 1/3 X +Zinf =0
'Dim a As Double, b As Double, c As Double, delta As Double
'a = 0.25
'b = (1 / 3)
'c = Zinf
'delta = b * b - 4 * a * c
'x1 = (-b + Sqr(delta)) / (2 * a)
'x2 = (-b - Sqr(delta)) / (2 * a)
'X = x1 * (F0 ^ 0.8 * h0)
'x0 = X
''x0 = ebc.Long
'
'dx = x0 / 10
'
'
'For i = 1 To 11
'x1 = dx * (i - 1)
'xy(i, 1) = x1
'X = x1 / (F0 ^ 0.8 * h0)
'Y = -(1 / 3 * h0) * X - 0.25 * h0 * X ^ 2
'xy(i, 2) = Y
'Next
'xam = ebc.tron_amo.Absava
'yam = ebc.tron_amo.radava
'x0 = xam
'y0 = yam
'For i = 1 To 11
'    X = xam + xy(i, 1)
'    Y = yam + xy(i, 2)
'    uc_g.dess_lign x0, y0, X, Y, couleur.bleu, 2
'    x0 = X
'    y0 = Y
'Next
'uc_g.dess_lign x0, y0, ebc.tron_amo.Absava + ebc.Long, ebc.tron_ava.radamo + h2, couleur.bleu, 2
'dess_pompe1 = True
'
'xam = ebc.tron_amo.Absamo
'yam = ebc.tron_amo.radamo + h1
'If le < ebc.tron_amo.Absava Then
'    xav = ebc.tron_amo.Absava - le
'    yav = ebc.tron_amo.radava + h1 + le * ebc.tron_amo.conduit.pente
'    uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
'    xam = xav
'    yam = yav
'End If
''dessin de la courbe
'dx = -le / 10
'If le > 5 Then dx = -5# / 10
'For i = 1 To 11
'x1 = dx * (i - 1)
'xy(i, 1) = x1
'X = x1 / le
'Y = 1 - (1 + X) ^ 2
'xy(i, 2) = Y * (h1 - he) + he - x1 * ebc.tron_amo.conduit.pente
'Next
'For i = 11 To 1 Step -1
'xav = ebc.tron_amo.Absava + xy(i, 1)
'yav = ebc.tron_amo.radava + xy(i, 2)
'    uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
'    xam = xav
'    yam = yav
'Next
'
'xav = ebc.tron_amo.Absava
'yav = ebc.tron_amo.radava + he
'uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
'xam = ebc.tron_amo.Absava + ebc.Long
'yam = ebc.tron_ava.radamo + h2
'xav = xam + ebc.tron_ava.conduit.Longueur
'yav = ebc.tron_ava.radava + h2
'uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
'
'
End Function

Function perte_charge_lin(vit As Double, dmm As Double, krugo As Double) As Double

Dim pas As Double, pasinf As Double, re As Double, coley As Double, xxx As Double, colez As Double, colex As Double
Dim cole1 As Integer, cole2 As Integer, jmpkm As Double, lambda As Double
Dim Nu As Double
Dim g As Double
g = 9.810001
Nu = 0.000001301
pas = 0.1
pasinf = 0.01
re = vit * dmm / 1000 / Nu
coley = 1000 * krugo / 3.71 / dmm
colez = 2.51 / re
Dim bok As Boolean
bok = True
xxx = pasinf
lambda = 0
While bok And xxx <= 1
'For xxx = pasinf To 1 Step pas
    colex = 1 / (xxx) ^ 0.5
    cole1 = Int(1000! * colex)
    COLEUN = 100000000# * colex
    cole2 = Int(1000! * (-2 * Log(coley + colez * colex) / Log(10)))
    COLEDE = 100000000# * (-2 * Log(coley + colez * colex) / Log(10))
'    Debug.Print "COEFFICIENT DE PERTE DE CHARGE "; xxx; COLEUN; COLEDE
    If cole1 = cole2 Then
        lambda = xxx
        bok = False
    ElseIf cole1 < cole2 Then
'    If cole1 <= cole2 Then
        pasinf = xxx - pas
        xxx = pasinf
        pas = pas / 10
    End If
    xxx = xxx + pas
Wend
If bok = True Then
MsgBox "erreur", vbExclamation, "perte cherge linéaire"
End If
'Next xxx
jmpkm = 500000! * lambda * vit ^ 2 / g / dmm
perte_charge_lin = jmpkm

End Function
Function rech_rugo(nat As String) As Double
Dim krugo As Double
Select Case nat
    Case "FONTE"
        krugo = 0.0001
    Case "PEHD "
        krugo = 0.000018
    Case "PVC  "
        krugo = 0.000018
    Case "ACIER"
        krugo = 0.000275
    Case Else
       krugo = 0.000018
End Select
rech_rugo = krugo
End Function
Function rech_Kbelier(nat As String) As Double
Dim kmate As Double
Select Case nat
    Case "FONTE"
       kmate = 0.6
    Case "PEHD "
       kmate = 83
    Case "PVC  "
       kmate = 33
    Case "ACIER"
       kmate = 0.5
    Case Else
       kmate = 83
End Select
rech_Kbelier = kmate
End Function

