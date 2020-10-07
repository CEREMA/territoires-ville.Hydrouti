Attribute VB_Name = "Module_Rural"

Sub Hye_Hyd_1r(l_int, fpas, k, s, c, Seuil, ByRef l() As Variant)
'Dim l() As Variant
Dim q1, edk As Double
Dim i As Integer
ReDim l(0)
Call AADD(l, 0#)
If k > 0 Then
    edk = Exp(-fpas / k)
Else
    edk = 0
End If
'Debug.Print UBound(l_int)
q1 = (1# - edk) * l_int(1) * s * c / 360#
Call AADD(l, q1)
For i = 2 To UBound(l_int)
    q1 = edk * q1 + (1# - edk) * l_int(i) * s * c / 360#
    Call AADD(l, q1)
Next
While q1 > Seuil / 1000#
    q1 = edk * q1
    Call AADD(l, q1)
Wend
'ReDim l_int(UBound(l))
'For i = 1 To UBound(l)
'    l_int(i) = l(i)
'Next
'Hye_Hyd_1r = l
End Sub
Function Perte_init(ByVal perte, ByRef l_int, ByVal fpas)
Dim i As Integer
Dim hp As Double

i = 0
'While perte > 0 And i < UBound(Q)
While perte > 0 And i < UBound(l_int)
    i = i + 1
    hp = l_int(i) * fpas / 60#
    If hp < perte Then
        l_int(i) = 0
        perte = perte - hp
    Else
        l_int(i) = (hp - perte) / fpas * 60#
        perte = 0
    End If
Wend
' Return(l_int)
End Function
Function Calcul_Infiltr(ByRef l_int, ByVal fpas, ByVal a, ByVal b, ByVal fc)
Dim hinf0, hp, tau, hinf As Double
Dim i As Integer
hinf0 = 0
tau = 0
For i = 1 To UBound(l_int)
    hp = l_int(i) * fpas / 60#
    hinf = h_tau(tau + fpas, a, b, fc) - hinf0
    If hinf <= hp Then
        hp = hp - hinf
        tau = tau + fpas
        hinf0 = hinf0 + hinf
    Else
        If hp > 0# Then
            tau = Itere_dtau(hp, tau, fpas, hinf0, a, b, fc, 0.01)
        End If
        hinf0 = hinf0 + hp
        hp = 0
    End If
    l_int(i) = hp / fpas * 60#
Next
'Return(l_int)
End Function

Function Itere_dtau(hp, t0, p, h0, a, b, fc, eps) As Double
Dim dt, HT, t As Double
dt = p / 2#
t = t0 + dt
HT = h_tau(t, a, b, fc) - h0
While Abs(HT - hp) > eps And dt > eps
    dt = dt / 2#
    If HT - hp > 0 Then
        t = t - dt
    Else
        t = t + dt
    End If
    HT = h_tau(t, a, b, fc) - h0
Wend
Itere_dtau = t
'Return(t)
End Function

Function h_tau(t, a, b, fc) As Double
Dim t0, b0 As Double
t0 = t / 60#
b0 = b / 60#
'b = b0
h_tau = fc * (t0 - a * b0 * (Exp(-t0 / b0) - 1))
End Function
Function Hye_Hyd_2r(l_int, fpas, k, s, Seuil)
Dim l() As Variant
Dim q1, q2, edk As Double
Dim i As Integer
ReDim l(0)
Call AADD(l, 0#)
If k > 0 Then
    edk = Exp(-fpas / k)
Else
    edk = 0
End If
q1 = (1# - edk) * l_int(1) * s / 360
q2 = (1# - edk) * q1
Call AADD(l, q2)
For i = 2 To UBound(l_int)
    q1 = edk * q1 + (1# - edk) * l_int(i) * s / 360
    q2 = edk * q2 + (1# - edk) * q1
    Call AADD(l, q2)
Next
While q2 > Seuil / 1000#
    q1 = edk * q1
    q2 = edk * q2 + (1# - edk) * q1
    Call AADD(l, q2)
Wend
ReDim l_int(UBound(l))
For i = 1 To UBound(l)
    l_int(i) = l(i)
Next
'Return(l)
End Function
Function Discret_hyeto(tpl, dt, HT, DM, HM, tp, fpas, l_int, ll) As Variant
Dim a, intmax, ia, t, t0, T1, T2, i1, i2, x1, x2, x3, x4, y1, y2, y3, y4 As Double
Dim l(10, 2) As Double
Dim i() As Variant
Dim ind, k As Integer
ReDim i(0)
If tpl = 3 Then
     dt = 0
     ind = 1
     While ind < UBound(l_int)
        dt = dt + l_int(ind)(1)
        ind = ind + 1
     Wend
Else
    T1 = tp - DM / 2#
    T2 = tp + DM / 2#
    If dt = DM Then
        i1 = 0
        T1 = 0
        i2 = 120 * (HT / dt)
    Else
    i1 = 120# * (HT - HM) / (dt - DM)
'    i1 = 120# * division((HT - HM), (dt - DM))
    i2 = 120 * (HM * (dt - DM) - DM * (HT - HM)) / (DM * (dt - DM))
'    i2 = 120 * division((HM * (dt - DM) - DM * (HT - HM)), (DM * (dt - DM)))
    x1 = i1 / T1
'    x1 = division(i1, T1)
    x2 = (i2 - i1) / (tp - T1)
    y2 = (tp * i1 - T1 * i2) / (tp - T1)
    x3 = (i2 - i1) / (tp - T2)
    y3 = (tp * i1 - T2 * i2) / (tp - T2)
    x4 = i1 / (T2 - dt)
    y4 = i1 * dt / (dt - T2)
    End If
    l(1, 1) = 0: l(1, 2) = 0
    l(2, 1) = T1: l(2, 2) = i1
    l(3, 1) = tp: l(3, 2) = i2
    l(4, 1) = T2: l(4, 2) = i1
    l(5, 1) = dt: l(5, 2) = 0
End If
If tpl = 2 Then
    ind = 2
    t = 0#
    t0 = 0#
    T1 = 0#
    T2 = 0#
    ReDim i(0)
    T2 = l(ind, 1)
    i2 = l(ind, 2)
    T1 = l(ind - 1, 1)
    i1 = l(ind - 1, 2)
    t = t + fpas
    While t <= dt
        If t <= T2 Then
            a = (((i2 - i1) / (T2 - T1)) * ((t + t0) / 2 - T1) + i1) * (t - t0) / fpas
        Else
            ia = 0
            While T2 < t And T2 < dt
'                ia = ia + (((i2 - i1) / (T2 - T1)) * ((T2 + t0) / 2 - T1) + i1) * (T2 - t0)
                ia = ia + ((division((i2 - i1), (T2 - T1))) * ((T2 + t0) / 2 - T1) + i1) * (T2 - t0)
                T1 = T2
                i1 = i2
                t0 = T2
                ind = ind + 1
                T2 = l(ind, 1)
                i2 = l(ind, 2)
            Wend
            ia = ia + (((i2 - i1) / (T2 - T1)) * ((t + t0) / 2 - T1) + i1) * (t - t0)
            a = ia / fpas
        End If
        If T2 = t And t < dt Then
            ind = ind + 1
            T1 = T2
            i1 = i2
            T2 = l(ind, 1)
            i2 = l(ind, 2)
        End If
        Call AADD(i, a)
        t0 = t
        t = t + fpas
    Wend
Else
    ind = 1
    t = 0#
    T1 = 0#
    T2 = 0#
    ReDim i(0)
    T2 = T1 + l_int(ind)(1)
    i1 = l_int(ind)(2)
    t = t + fpas
If t >= dt Then
    a = i1 / fpas
        Call AADD(i, a)
Else
    While t <= dt
        If t <= T2 Then
            a = i1
        Else
            ia = 0
            While T2 < t And T2 < dt
                ia = i1 * (T2 - T1) + ia
                T1 = T2
                ind = ind + 1
                T2 = T1 + l_int(ind)(1)
                i1 = l_int(ind)(2)
            Wend
            ia = i1 * (t - T1) + ia
            a = ia / fpas
        End If
        If T2 = t And t < dt Then
            ind = ind + 1
            T1 = T2
            T2 = T1 + l_int(ind)(1)
            i1 = l_int(ind)(2)
        End If
        Call AADD(i, a)
        T1 = Maxi(T1, t)
        t = t + fpas
    Wend
End If
End If
ReDim ll(UBound(i))
For k = 1 To UBound(i)
    ll(k) = i(k)
Next
Discret_hyeto = i
End Function


Sub AADD(ByRef tabl() As Variant, Q As Variant)    'ajout d'une valeur à une variable tableau
Dim n As Integer
DoEvents
n = UBound(tabl) + 1
ReDim Preserve tabl(n)
tabl(n) = Q
End Sub
Function division(ByVal a As Variant, ByVal b As Variant) As Double
If b = 0 Then
 division = 0
Else
 division = a / b
End If

End Function
Function MAx_listeh(ByRef l) As Double
Dim i As Integer
Dim m As Double
m = l(2)
For i = 3 To UBound(l)
    m = Maxi(m, l(i))
Next
X = m
MAx_listeh = m
End Function
Function Maxi(a, b)
If a >= b Then
    Maxi = a
Else
    Maxi = b
End If
End Function
Function Mini(a, b)
If a <= b Then
    Mini = a
Else
    Mini = b
End If
End Function
Function MAxh_listeh(ByRef l) As Integer
Dim i, tm As Integer
Dim m As Double
m = l(2)
tm = 2
For i = 3 To UBound(l)
    If m < l(i) Then
        m = l(i)
        tm = i
    End If
Next
tm = tm - 1
MAxh_listeh = tm
End Function
Function Som_listeh(ByRef l) As Double
Dim i As Integer
Dim m As Double
m = 0#
For i = 2 To UBound(l)
    m = m + l(i)
Next
Som_listeh = m
End Function

