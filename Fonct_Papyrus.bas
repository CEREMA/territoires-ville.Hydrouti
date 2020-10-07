Attribute VB_Name = "Fonct_Papyrus"
Type deb_vit
    debit As Double
    vitesse As Double
    End Type
Type per_sur
    per As Double
    sur As Double
    End Type
Type conduite_papy
    nom As String
    Longueur As Double
    Diametre As Double
    pente As Double
    rugosite As Integer
    largfond As Double
    larghaut As Double
    ouvert As Integer
    typ As Double
    symetrique As String
    materiau As String
    epaisseur As Double
    nomtable As String
    
    End Type
Type conduite
    Longueur As Double
    Diametre As Double
    pente As Double
    rugosite As Integer
    typ As Double
    End Type
    
Sub cana(ByRef X As conduite, lc As Variant)
'As Double
ReDim lc(7)
lc(1) = X.Longueur
lc(2) = X.Diametre
lc(3) = X.pente
lc(4) = X.rugosite
lc(6) = 0  ' X.ouvert  ' =0
lc(7) = X.typ ' =2
'cana = lc
End Sub

Function debit_ps(ByRef lo As conduite) As Double
Dim rh, ray, kq, v, s, q As Double
Dim lps As per_sur
lps = car_ps(lo)
If lps.per = 0 Then
    q = 0
Else
    p = lps.per
    s = lps.sur
    kg = lo.rugosite * Sqr(Maxi(lo.pente, 0))
    rh = s / p
    v = kg * rh ^ (2# / 3#)
    q = v * s
 '   Debug.Print v, q
End If
debit_ps = q
End Function
Function debvit_ps(ByRef lo As conduite) As deb_vit
Dim rh, ray, kq, v, s, q As Double
Dim lps As per_sur
lps = car_ps(lo)
If lps.per = 0 Then
    q = 0
    v = 0
Else
    p = lps.per
    s = lps.sur
    kg = lo.rugosite * Sqr(Maxi(lo.pente, 0))
    rh = s / p
    v = kg * rh ^ (2# / 3#)
    q = v * s
End If
With debvit_ps
    .debit = q
    .vitesse = v
End With
End Function


Function car_ps(l As conduite) As per_sur

With l
    Select Case .typ
        Case 2
            car_ps = dps_cir(l)
'         Case 3
'            car_ps = dps_rct(l)
'        Case 4, 16
'            car_ps = dps_tra(l)
'        Case 13
'            car_ps = dps_rct(l)
'        Case 5, 9
'            car_ps = dps_ovo(l)
'        Case 18
'            car_ps = dps_oau(l)
        Case Else
    End Select
End With
End Function
Function dps_cir(l As conduite) As per_sur
Dim ray, p, s As Double
With l
    ray = .Diametre / 2#
    p = 3.14159 * .Diametre
    s = 3.14159 * .Diametre ^ 2 * 0.25
End With
With dps_cir
    .per = p
    .sur = s
End With
End Function
'
'
'                ltc = calc_par(canax)
'                qvm = caltran1(qcal, ct, ltc)
''*         ? "qvm ",qvm(1),qvm(2),qvm(3),qvm(4),qvm(5)
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
'
Function caltran1(a, b, ltc As Variant) As Variant
Dim lqv(5) As Double
Dim i As Integer
Dim q, lm, haut, pen, ray, Ks, kg, lkg, louvr, epsi As Double
louvr = b(1)
ray = b(2) / 2#
pen = b(3)
Ks = b(4)
epsi = 0.005
kg = Ks * Sqr(pen)
lkg = kg * ray ^ (8# / 3#)
q = a / (1000# * lkg)
i = 1
While i < UBound(ltc) And q >= ltc(Mini(i, UBound(ltc)), 1)
    i = i + 1
Wend
i = i - 1
If i < 1 Then
i = 1
End If
If i < UBound(ltc) Then
    alpha = (q - ltc(i, 1)) / (ltc(i + 1, 1) - ltc(i, 1))
Else
    alpha = 1
    i = i - 1
End If
ch = lkg / (ray * ray) * (alpha * (ltc(i + 1, 3) - ltc(i, 3)) + ltc(i, 3))
lm = ray * (alpha * (ltc(i + 1, 4) - ltc(i, 4)) + ltc(i, 4))
haut = ray * (alpha * (ltc(i + 1, 5) - ltc(i, 5)) + ltc(i, 5))
'debug.print"Debit =",q," Vs= ",Vs," DT =",Louvr/Vs," s","lm = ",lm,"haut = ",haut
vs = kg * ray ^ (2# / 3#) * (alpha * (ltc(i + 1, 2) - ltc(i, 2)) + ltc(i, 2))
' Debug.Print "Debit =", a, " Vs= ", Vs, " DT =", louvr / Vs, " s Celerite =", ch, "m/s"
lqv(1) = a / 1000#
lqv(2) = vs
lqv(3) = ch
lqv(4) = lm
lqv(5) = haut
'return(lqv)
caltran1 = lqv
End Function
Function calc_par(lo As conduite) As Variant
Dim l() As Variant
Dim X As Integer
ReDim l(51, 5)
With lo
    Select Case .typ
        Case 2
            X = calc_cir(lo, l)
'        Case 9
'            X = calc_ovo(lo, l)
'        Case 16
'            X = calc_tra(lo, l)
'        Case 13
'            X = calc_rct(lo, l)
'        Case 18
'            X = calc_oau(lo, l)
'        Case 7
'            X = calc_do(lo, l)
'        Case 8
'            X = calc_do(lo, l)
'        Case 14
'            X = calc_do(lo, l)
'        Case 15
'  '         l = calc_do(lo)
'        Case 19
'  '         l = calc_do(lo)
    End Select
    calc_par = l
End With
End Function
Function calc_cir(lo As conduite, lct()) As Integer
Dim lc(5) As Double
Dim i, j, k As Integer
Dim lqv(5) As Double
Dim pi, Sh, rh, Ph, haut, Ks, kg, lkg, alpha, dalpha, epsi As Double
Dim sina, sin2a, sa, qa, q0, dq0, dq, ds, ch As Double
Dim q, lm, louvr As Double
ReDim lct(101, 5) As Variant
pi = Atn(1#) * 4
alpha = 0#
dpha = pi / 100#
i = 1
q0 = 0
qa = 0#
s0 = 0#
sa = 0#
dq0 = 0
ch = 0
vs = 0#
ch = 0#
lm = 0#
haut = 0#

lc(1) = qa
lc(2) = vs
lc(3) = ch
lc(4) = lm
lc(5) = haut
For j = 1 To 5
    lct(i, j) = lc(j)
Next

alpha = alpha + dpha

For i = 2 To 101
 '  lc={}
    sina = Sin(alpha)
    sin2a = Sin(2 * alpha)
    sa = (alpha - sin2a / 2)
    qa = ((alpha - (sin2a / 2)) ^ (5# / 3#)) * ((2 * alpha) ^ (-2# / 3#))
    dq0 = qa - q0
    dq = 5# / 3# * (alpha - (sin2a / 2)) ^ (2# / 3#) * (2 * alpha) ^ (-2# / 3#) * 2 * sina ^ 2
    dq = dq + (alpha - (sin2a / 2)) ^ (5# / 3#) * (-4# / 3#) * (2 * alpha) ^ (-5# / 3#)
    ds = sa - s0
    ds = 2 * sina ^ 2
    If ds = 0 Then
        ch = 0
    Else
        ch = dq / ds
    End If
    q0 = qa
    s0 = sa
    lm = sina * 2#
    haut = 1# - Sin((pi / 2) - alpha)
    Sh = (alpha - sin2a / 2)
    Ph = 2 * alpha
    rh = Sh / Ph
    vs = rh ^ (2# / 3#)
   
    lc(1) = qa
    lc(2) = vs
    lc(3) = ch
    lc(4) = lm
    lc(5) = haut
 '      Debug.Print i; 180 * alpha / pi; lc(1); lc(2); lc(3); lc(4); lc(5)
    For j = 1 To 5
        lct(i, j) = lc(j)
    Next
' Debug.Print "message," + Str(i) + Str(lct(i, 1)) + " " + Str(lct(i, 2)) + " " + Str(lct(i, 3)) + " " + Str(lct(i, 4)) + " " + Str(lct(i, 5))
    
    alpha = alpha + dpha
Next
l = lct
calc_cir = i
'return(lct)
End Function
Function pent_mot0(ByRef lo As conduite, ByVal q As Double) As Double
Dim lps As per_sur
Dim p, s, kg, rh, Im As Double
lps = car_ps(lo)
p = lps.per
s = lps.sur
kg = lo.rugosite
If p = 0 Then
rh = 0
Im = 0
Else
rh = s / p
Im = (q / (kg * s * rh ^ (2# / 3#))) ^ 2
End If
'  Debug.Print "kg,ray,rh,s,q,im*100", kg, ray, Rh, s, q, im * 100
pent_mot0 = Im
End Function
Public Function inters(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double, ByRef xi As Double, ByRef yi As Double)
Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
'Dim xi As Double, yi As Double
',x1,y1,x2,y2,x3,y3,x4,y4,p[2]


'x1 = p1[1]
'y1 = p1[2]
'x2 = p2[1]
'y2 = p2[2]
'x3 = p3[1]
'y3 = p3[2]
'x4 = p4[1]
'y4 = p4[2]

If x2 = x1 Then
        xi = x1
        a2 = (y4 - y3) / (x4 - x3)
        b2 = y3 - a2 * x3
        yi = a2 * x1 + b2
ElseIf x3 = x4 Then
        xi = x3
        a1 = (y2 - y1) / (x2 - x1)
        b1 = y1 - a1 * x1
        yi = a1 * x3 + b1
Else
        a1 = (y2 - y1) / (x2 - x1)
        b1 = y1 - a1 * x1
        a2 = (y4 - y3) / (x4 - x3)
        b2 = y3 - a2 * x3
        If a2 - a1 = 0 Then
        xi = x1
        yi = y1
        Else
        xi = (b1 - b2) / (a2 - a1)
        yi = b1 + a1 * xi
        End If
End If

'Return ( p )
inters = True
End Function
