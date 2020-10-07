Attribute VB_Name = "Module2piezo"
Public Function inter_piezo_eau(ByRef tr As troncon, ByRef tr_res As debit_conduit)
Dim xam As Double, yam As Double, xav As Double, yav As Double, pcana As Double
Dim xamp As Double, yamp As Double, xavp As Double, yavp As Double, ppiezo As Double
Dim xi As Double, yi As Double
Dim p As points, p_int As points
Dim ok As Boolean
' intersection avec la generaatrice superieure
    ppiezo = tr_res.pentemotrice
xam = tr.Absamo
yam = tr.radamo + tr.conduit.Diametre
xav = tr.Absava
yav = tr.radava + tr.conduit.Diametre
xavp = tr.Absava
yavp = maximum(tr_res.zphe_ava, tr.radava + tr_res.hauteur)
tr_res.piezoava = yavp
tr_res.zeau_amo.X = xam
tr_res.zeau_ava.X = xav
tr_res.zeau_ava.Y = yavp
If yavp >= yav Then
    tr_res.zeau_ava.Y = yav
    xamp = tr.Absamo
    yamp = yavp + ((tr.conduit.Longueur) * ppiezo)
    If yamp >= yam Then
           tr_res.piezoamo = yamp
            tr_res.zeau_amo.Y = yam
            tr_res.piezointer0.X = xav
            tr_res.piezointer0.Y = tr_res.piezoava
            tr_res.p_Eau_inter0 = tr_res.zeau_ava
            tr_res.piezointer.X = xam
            tr_res.piezointer.Y = tr_res.piezoamo
            tr_res.p_Eau_inter = tr_res.zeau_amo
            tr_res.p_Eau_inter1 = tr_res.p_Eau_inter0
            tr_res.p_Eau_inter2 = tr_res.p_Eau_inter
    Else
        ok = inters(xam, yam, xav, yav, xamp, yamp, xavp, yavp, xi, yi)
        tr_res.piezointer0.X = xi
        tr_res.piezointer0.Y = yi
        tr_res.p_Eau_inter0.X = xi
        tr_res.p_Eau_inter0.Y = yi
        'recherche pt de contact avec ligne d'eau
        p_int = rech_inter(xi, yi, tr, tr_res)
        tr_res.piezointer.X = p_int.X
        tr_res.piezointer.Y = p_int.Y
        tr_res.p_Eau_inter.X = p_int.X
        tr_res.p_Eau_inter.Y = p_int.Y
        If p_int.X > tr.Absamo Then
            tr_res.piezoamo = tr_res.hauteur + tr.radamo
            tr_res.zeau_amo.Y = tr_res.piezoamo
        Else
            tr_res.piezoamo = p_int.Y
            tr_res.zeau_amo.Y = p_int.Y
        
        End If
    
    End If
 Else
    tr_res.zeau_ava.Y = yavp
    tr_res.piezointer0.X = xav
    tr_res.piezointer0.Y = tr_res.piezoava
    tr_res.p_Eau_inter0 = tr_res.zeau_ava
        'recherche pt de contact avec ligne d'eau
    p_int = rech_inter(xav, yavp, tr, tr_res)
    tr_res.piezointer.X = p_int.X
    tr_res.piezointer.Y = p_int.Y
    tr_res.p_Eau_inter.X = p_int.X
    tr_res.p_Eau_inter.Y = p_int.Y
    ' a recalculer
     If p_int.X > tr.Absamo Then
        tr_res.piezoamo = tr_res.hauteur + tr.radamo
        tr_res.zeau_amo.Y = tr_res.piezoamo
    Else
        tr_res.piezoamo = p_int.Y
        tr_res.zeau_amo.Y = p_int.Y
    
    End If
End If

GoTo suite



tr_res.zeau_amo.X = xam
tr_res.zeau_ava.X = xav
tr_res.piezoava = yavp
If yamp <= yam Then
    tr_res.piezoamo = yam
    tr_res.zeau_amo.Y = yam
Else
    tr_res.piezoamo = yamp
    If yamp > (tr.conduit.Diametre + tr.radamo) Then
        tr_res.zeau_amo.Y = tr.conduit.Diametre + tr.radamo
        Else
        tr_res.zeau_amo.Y = yamp
    End If
End If
If yavp > (tr.conduit.Diametre + tr.radava) Then
    tr_res.zeau_ava.Y = tr.conduit.Diametre + tr.radava
ElseIf yavp > yav Then
        tr_res.zeau_ava.Y = yavp
Else
    tr_res.zeau_ava.Y = yav
End If

suite:
tr_res.hautamo = tr_res.zeau_amo.Y - tr.radamo
tr_res.hautava = tr_res.zeau_ava.Y - tr.radava


End Function
Function rech_inter(ByVal xi As Double, ByVal yi, ByRef tr As troncon, ByRef tr_res As debit_conduit) As points
Dim X As Double, Y As Double, h As Double, hei As Double, heau As Double, pente As Double
Dim xam As Double, yam As Double, xav As Double, yav As Double, pcana As Double, ppiezo As Double
Dim dx As Double, dh As Double, hei0 As Double, hf As Double
Dim pi As points
heau = tr_res.hauteur
pcana = tr.conduit.pente
X = xi: Y = yi
xav = tr.Absava: yav = tr.radava
xam = tr.Absamo: yam = tr.radamo
hei = yi - (xav - xi) * pcana - yav
hei0 = hei
'ppiezo = pent_mot0(tr.conduit, tr_res.debit)
ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
dh = 0.01
While hei > heau And X > xam
    ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
    If (hei - dh) <= heau Then
        dh = hei - heau
    End If
    dx = -dh / (ppiezo - pcana)
    If (X - dx) < xam Then
        dx = X - xam
       dh = dx * (ppiezo - pcana)
    End If
    hei = hei - dh
    X = X - dx
Wend

rech_inter.X = X
rech_inter.Y = (xav - X) * pcana + tr.radava + hei


'calcul des pts inter 1 et 2
hf = hei
X = xi: Y = yi
hei = hei0
'ppiezo = pent_mot0(tr.conduit, tr_res.debit)
ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
heau = (hei - hf) * 0.6 + hf
dh = 0.01
While hei > heau And X > xam
    ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
    If (hei - dh) <= heau Then
        dh = hei - heau
    End If
    dx = -dh / (ppiezo - pcana)
    If (X - dx) < xam Then
        dx = X - xam
       dh = dx * (ppiezo - pcana)
    End If
    hei = hei - dh
    X = X - dx
Wend
tr_res.p_Eau_inter1.X = X
tr_res.p_Eau_inter1.Y = (xav - X) * pcana + tr.radava + hei
tr_res.piezointer1.X = X
tr_res.piezointer1.Y = (xav - X) * pcana + tr.radava + hei



X = xi: Y = yi
hei = hei0
'ppiezo = pent_mot0(tr.conduit, tr_res.debit)
ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
heau = (hei - hf) * 0.4 + hf
dh = 0.01
While hei > heau And X > xam
    ppiezo = pent_mot_h(hei, tr.conduit, tr_res.debit)
    If (hei - dh) <= heau Then
        dh = hei - heau
    End If
    dx = -dh / (ppiezo - pcana)
    If (X - dx) < xam Then
        dx = X - xam
       dh = dx * (ppiezo - pcana)
    End If
    hei = hei - dh
    X = X - dx
Wend
tr_res.p_Eau_inter2.X = X
tr_res.p_Eau_inter2.Y = (xav - X) * pcana + tr.radava + hei
tr_res.piezointer2.X = X
tr_res.piezointer2.Y = (xav - X) * pcana + tr.radava + hei

End Function
Function pent_mot_h(ByVal h As Double, ByRef co As conduite, ByVal Q As Double) As Double
Dim beta As Double, alpha As Double
Dim s As Double, p As Double, v As Double, zhe As Double

Dim rh As Double, Im As Double
    If h / co.Diametre < 1 Then
        beta = 2 * arccosinus((1 - 2 * h / co.Diametre))
    Else
        beta = 2 * pi
    End If
    alpha = beta / 2#
    s = (co.Diametre ^ 2) * (alpha - Sin(beta) / 2) / 4
    p = co.Diametre * alpha
    kg = co.rugosite
If p = 0 Then
rh = 0
Im = 0
Else
    rh = s / p
    Im = (Q / (kg * s * rh ^ (2# / 3#))) ^ 2
End If
'  Debug.Print "kg,ray,rh,s,q,im*100", kg, ray, Rh, s, q, im * 100
pent_mot_h = Im
End Function
Function long_chute(ByRef ebc As st_Chute, ByRef res_am As debit_conduit, ByRef res_av As debit_conduit, ByRef hee As Double, ByRef lee As Double) As Double
Dim dh As Double, h0 As Double, h1 As Double, h2 As Double, D0 As Double
D0 = ebc.tron_amo.conduit.Diametre
h0 = res_am.hautava
h2 = res_av.hautamo
dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
long_chute = dh
Dim F0 As Double, he As Double
F0 = res_am.debit / ((h0 ^ 2) * ((9.81 * D0) ^ 0.5))
he = h0 * ((2 * F0 ^ 2) / (1 + 2 * F0 ^ 2)) ^ (2# / 3#)
Dim le As Double
le = h0 * (5 + 0.9 * F0)
lee = le
hee = he
Dim Zsup1 As Double, Zinf As Double
dh = dh - h2
Zinf = -dh / h0
dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
dh = dh + he - h2
ebc.h0 = dh
Zsup1 = -dh / h0
Zsup1 = Zinf
' Zsup1=-1/3 X -1/4 X^2
'1/4 X^2 + 1/3 X +Zinf =0
Dim a As Double, b As Double, c As Double, delta As Double
a = 0.25
b = (1 / 3)
c = Zinf
delta = b * b - 4 * a * c
x1 = (-b + Sqr(delta)) / (2 * a)
x2 = (-b - Sqr(delta)) / (2 * a)
Dim X As Double
X = x1 * (F0 ^ 0.8 * h0)
' Zsup1=1 -(1/3-0.06)X -1/4 X^2
'1/4 X^2 +(1/3-0.06) X +Zsup -1 =0
'Dim a As Double, b As Double, c As Double, delta As Double
a = 0.25
b = (1 / 3 - 0.06 * he / h0)
c = Zsup1 - he / h0
delta = b * b - 4 * a * c
x1 = (-b + Sqr(delta)) / (2 * a)
x2 = (-b - Sqr(delta)) / (2 * a)
'Dim X As Double
X = x1 * (F0 ^ 0.8 * h0)
long_chute = X
End Function

Function dess_chute1(ByRef uc_g As UC_graphique, ByRef ebc As st_Chute, ByRef res_am As debit_conduit, ByRef res_av As debit_conduit) As Boolean
Dim dh As Double, h0 As Double, h1 As Double, h2 As Double, D0 As Double

Dim xy(11, 2) As Double, X As Double, dx As Double, Y As Double, x0 As Double, y0 As Double
uc_g.redef_drwidth 2

D0 = ebc.tron_amo.conduit.Diametre
h0 = res_am.hautava
h2 = res_av.hautamo
h1 = res_am.hautava

dh = ebc.tron_amo.radava - ebc.tron_ava.radamo
Dim F0 As Double, he As Double
F0 = res_am.debit / ((h0 ^ 2) * ((9.81 * D0) ^ 0.5))
he = h0 * ((2 * F0 ^ 2) / (1 + 2 * F0 ^ 2)) ^ (2# / 3#)
Dim le As Double
le = h0 * (5 + 0.9 * F0)
Dim Zsup1 As Double, Zinf As Double

x0 = ebc.Long

dx = x0 / 10
For i = 1 To 11
x1 = dx * (i - 1)
xy(i, 1) = x1
X = x1 / (F0 ^ 0.8 * h0)
Y = he - (1 / 3 * h0 - 0.06 * he) * X - 0.25 * h0 * X ^ 2
xy(i, 2) = Y
Next
xam = ebc.tron_amo.Absava
yam = ebc.tron_amo.radava
x0 = xam
y0 = yam + he
For i = 1 To 11
    X = xam + xy(i, 1)
    Y = yam + xy(i, 2)
    uc_g.dess_lign x0, y0, X, Y, couleur.bleu, 2
    x0 = X
    y0 = Y
Next
'courbe inferieure
dh = dh - h2
Zinf = -dh / h0
' Zsup1=-1/3 X -1/4 X^2
'1/4 X^2 + 1/3 X +Zinf =0
Dim a As Double, b As Double, c As Double, delta As Double
a = 0.25
b = (1 / 3)
c = Zinf
delta = b * b - 4 * a * c
x1 = (-b + Sqr(delta)) / (2 * a)
x2 = (-b - Sqr(delta)) / (2 * a)
X = x1 * (F0 ^ 0.8 * h0)
x0 = X
'x0 = ebc.Long

dx = x0 / 10


For i = 1 To 11
x1 = dx * (i - 1)
xy(i, 1) = x1
X = x1 / (F0 ^ 0.8 * h0)
Y = -(1 / 3 * h0) * X - 0.25 * h0 * X ^ 2
xy(i, 2) = Y
Next
xam = ebc.tron_amo.Absava
yam = ebc.tron_amo.radava
x0 = xam
y0 = yam
For i = 1 To 11
    X = xam + xy(i, 1)
    Y = yam + xy(i, 2)
    uc_g.dess_lign x0, y0, X, Y, couleur.bleu, 2
    x0 = X
    y0 = Y
Next
uc_g.dess_lign x0, y0, ebc.tron_amo.Absava + ebc.Long, ebc.tron_ava.radamo + h2, couleur.bleu, 2
dess_chute1 = True

xam = ebc.tron_amo.Absamo
yam = ebc.tron_amo.radamo + h1
If le < ebc.tron_amo.Absava Then
    xav = ebc.tron_amo.Absava - le
    yav = ebc.tron_amo.radava + h1 + le * ebc.tron_amo.conduit.pente
    uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
    xam = xav
    yam = yav
End If
'dessin de la courbe
dx = -le / 10
If le > 5 Then dx = -5# / 10
For i = 1 To 11
x1 = dx * (i - 1)
xy(i, 1) = x1
X = x1 / le
Y = 1 - (1 + X) ^ 2
xy(i, 2) = Y * (h1 - he) + he - x1 * ebc.tron_amo.conduit.pente
Next
For i = 11 To 1 Step -1
xav = ebc.tron_amo.Absava + xy(i, 1)
yav = ebc.tron_amo.radava + xy(i, 2)
    uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
    xam = xav
    yam = yav
Next

xav = ebc.tron_amo.Absava
yav = ebc.tron_amo.radava + he
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2
xam = ebc.tron_amo.Absava + ebc.Long
yam = ebc.tron_ava.radamo + h2
xav = xam + ebc.tron_ava.conduit.Longueur
yav = ebc.tron_ava.radava + h2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 2


End Function



